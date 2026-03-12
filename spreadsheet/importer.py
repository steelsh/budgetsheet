"""
Excel Importer
==============
Reads an .xlsx file with openpyxl and populates the DB:
  - Parses cell values and formulas
  - Converts Excel formulas → Python expressions
  - Builds the dependency graph (CellDependency)

Usage:
    from spreadsheet.importer import import_excel
    sheet = import_excel('path/to/file.xlsx', sheet_index=0, sheet_name='Бюджет')
"""

import re
from .models import Sheet, Cell, CellDependency
from .formula_engine import formula_to_python, extract_cell_refs


def _get_cell_format(number_format: str):
    """Guess format_type from Excel number format string."""
    if not number_format:
        return '', 2
    nf = number_format.lower()
    if any(c in nf for c in ['₸', '$', '€', '£', '"руб"', 'rub', '#,##0']):
        return 'currency', 2
    if '%' in nf:
        return 'percent', 1
    if '0.00' in nf or '#,##' in nf:
        return 'number', 2
    if '0.0' in nf:
        return 'number', 1
    return '', 0


def _color_to_hex(color_obj):
    """Convert openpyxl Color to hex string."""
    if color_obj is None:
        return ''
    try:
        if hasattr(color_obj, 'rgb') and color_obj.rgb and color_obj.rgb != '00000000':
            rgb = color_obj.rgb
            # Skip fully transparent or default
            if rgb.startswith('FF') or rgb.startswith('ff'):
                return '#' + rgb[2:8]
            return '#' + rgb[2:8] if len(rgb) == 8 else ''
    except Exception:
        pass
    return ''


def import_excel(filepath: str, sheet_index: int = 0, sheet_name: str = None) -> Sheet:
    """
    Import an Excel sheet into the database.
    Returns the created Sheet instance.
    """
    try:
        from openpyxl import load_workbook
        from openpyxl.utils import get_column_letter
    except ImportError:
        raise ImportError("openpyxl is required: pip install openpyxl")

    wb = load_workbook(filepath, data_only=False)

    if sheet_name and sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.worksheets[sheet_index]

    # Determine actual used range
    max_row = ws.max_row or 1
    max_col = ws.max_column or 1

    # Create or replace sheet
    name = sheet_name or ws.title
    sheet, _ = Sheet.objects.get_or_create(name=name)
    # Clear existing data for re-import
    sheet.cells.all().delete()
    sheet.dependencies.all().delete()

    cells_to_create = []
    formula_cells = []

    # Track merged cells
    merged_ranges = {}
    for merge_range in ws.merged_cells.ranges:
        min_row, min_col = merge_range.min_row - 1, merge_range.min_col - 1
        max_row_m, max_col_m = merge_range.max_row - 1, merge_range.max_col - 1
        merged_ranges[(min_row, min_col)] = {
            'row_span': max_row_m - min_row + 1,
            'col_span': max_col_m - min_col + 1,
        }
        # Mark spanned cells to skip
        for r in range(min_row, max_row_m + 1):
            for c in range(min_col, max_col_m + 1):
                if not (r == min_row and c == min_col):
                    merged_ranges[(r, c)] = 'skip'

    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col)):
        for col_idx, excel_cell in enumerate(row):
            rc = (row_idx, col_idx)

            if merged_ranges.get(rc) == 'skip':
                continue

            merge_info = merged_ranges.get(rc, {})

            raw = excel_cell.value
            formula = None
            py_formula = None
            cell_type = 'input'
            is_editable = True

            if isinstance(raw, str) and raw.startswith('='):
                formula = raw
                py_formula = formula_to_python(raw[1:])
                cell_type = 'formula'
                is_editable = False
                value = '0'  # Will be recalculated
                formula_cells.append((row_idx, col_idx, raw))
            elif raw is None:
                value = ''
                cell_type = 'label'
                is_editable = False
            elif isinstance(raw, (int, float)):
                value = str(raw)
                is_editable = True
            elif isinstance(raw, str):
                value = raw
                # Detect headers/labels: non-numeric strings in first few rows or columns
                if col_idx < 3 or row_idx < 3:
                    cell_type = 'header' if row_idx < 2 else 'label'
                    is_editable = False
            else:
                value = str(raw) if raw is not None else ''

            # Styling
            font = excel_cell.font
            fill = excel_cell.fill
            fmt_type, dec = _get_cell_format(excel_cell.number_format)

            bold = font.bold if font else False
            italic = font.italic if font else False
            text_color = ''
            bg_color = ''

            if font and font.color:
                text_color = _color_to_hex(font.color)
            if fill and fill.fgColor:
                bg_color = _color_to_hex(fill.fgColor)

            cell_obj = Cell(
                sheet=sheet,
                row=row_idx,
                col=col_idx,
                value=value,
                raw_value=str(raw) if raw is not None else '',
                formula=formula,
                python_formula=py_formula,
                cell_type=cell_type,
                is_editable=is_editable,
                format_type=fmt_type,
                decimal_places=dec,
                row_span=merge_info.get('row_span', 1),
                col_span=merge_info.get('col_span', 1),
                bold=bool(bold),
                italic=bool(italic),
                text_color=text_color,
                bg_color=bg_color,
                number_format=excel_cell.number_format or '',
            )
            cells_to_create.append(cell_obj)

    Cell.objects.bulk_create(cells_to_create, ignore_conflicts=True)

    # Build dependency graph
    deps_to_create = []
    for (row_idx, col_idx, formula) in formula_cells:
        refs = extract_cell_refs(formula)
        for (r, c) in refs:
            deps_to_create.append(CellDependency(
                sheet=sheet,
                source_row=r,
                source_col=c,
                target_row=row_idx,
                target_col=col_idx,
            ))

    if deps_to_create:
        CellDependency.objects.bulk_create(deps_to_create, ignore_conflicts=True)

    # Initial full recalculation
    _recalc_all(sheet)

    return sheet


def _recalc_all(sheet):
    """Do a full recalculation pass after import."""
    from .formula_engine import recalculate_sheet, make_eval_context

    changed, cell_values = recalculate_sheet(sheet)
    if changed:
        cells = list(sheet.cells.filter(cell_type='formula'))
        cell_map = {(c.row, c.col): c for c in cells}
        to_save = []
        for (r, c), val in changed.items():
            cell = cell_map.get((r, c))
            if cell:
                cell.value = val
                to_save.append(cell)
        if to_save:
            Cell.objects.bulk_update(to_save, ['value'])
