"""
Demo seeder — creates a realistic P&L / Budget financial table for demonstration.
Run with: python manage.py seed_demo
"""

from .models import Sheet, Cell, CellDependency
from .formula_engine import formula_to_python, extract_cell_refs


DEMO_DATA = [
    # Row 0: Title
    {'r': 0, 'c': 0, 'v': 'БЮДЖЕТ ДОХОДОВ И РАСХОДОВ', 'type': 'header', 'bold': True, 'col_span': 7},

    # Row 1: Column headers
    {'r': 1, 'c': 0, 'v': 'Статья', 'type': 'header', 'bold': True},
    {'r': 1, 'c': 1, 'v': 'Янв', 'type': 'header', 'bold': True},
    {'r': 1, 'c': 2, 'v': 'Фев', 'type': 'header', 'bold': True},
    {'r': 1, 'c': 3, 'v': 'Мар', 'type': 'header', 'bold': True},
    {'r': 1, 'c': 4, 'v': 'Апр', 'type': 'header', 'bold': True},
    {'r': 1, 'c': 5, 'v': 'Май', 'type': 'header', 'bold': True},
    {'r': 1, 'c': 6, 'v': 'Июн', 'type': 'header', 'bold': True},
    {'r': 1, 'c': 7, 'v': 'ИТОГО', 'type': 'header', 'bold': True},
    {'r': 1, 'c': 8, 'v': '% выполнения', 'type': 'header', 'bold': True},

    # Row 2: Section header - Revenue
    {'r': 2, 'c': 0, 'v': '─── ДОХОДЫ ───', 'type': 'label', 'bold': True, 'bg': '#e8f5e9'},

    # Row 3: Product sales
    {'r': 3, 'c': 0, 'v': 'Продажи продукта A', 'type': 'label'},
    {'r': 3, 'c': 1, 'v': '1200000', 'type': 'input', 'fmt': 'currency'},
    {'r': 3, 'c': 2, 'v': '1350000', 'type': 'input', 'fmt': 'currency'},
    {'r': 3, 'c': 3, 'v': '1500000', 'type': 'input', 'fmt': 'currency'},
    {'r': 3, 'c': 4, 'v': '1450000', 'type': 'input', 'fmt': 'currency'},
    {'r': 3, 'c': 5, 'v': '1600000', 'type': 'input', 'fmt': 'currency'},
    {'r': 3, 'c': 6, 'v': '1750000', 'type': 'input', 'fmt': 'currency'},
    {'r': 3, 'c': 7, 'v': '=SUM(B4:G4)', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 3, 'c': 8, 'v': '=H4/9000000*100', 'type': 'formula', 'fmt': 'percent'},

    # Row 4: Service revenue
    {'r': 4, 'c': 0, 'v': 'Продажи продукта B', 'type': 'label'},
    {'r': 4, 'c': 1, 'v': '800000', 'type': 'input', 'fmt': 'currency'},
    {'r': 4, 'c': 2, 'v': '850000', 'type': 'input', 'fmt': 'currency'},
    {'r': 4, 'c': 3, 'v': '900000', 'type': 'input', 'fmt': 'currency'},
    {'r': 4, 'c': 4, 'v': '920000', 'type': 'input', 'fmt': 'currency'},
    {'r': 4, 'c': 5, 'v': '950000', 'type': 'input', 'fmt': 'currency'},
    {'r': 4, 'c': 6, 'v': '1000000', 'type': 'input', 'fmt': 'currency'},
    {'r': 4, 'c': 7, 'v': '=SUM(B5:G5)', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 4, 'c': 8, 'v': '=H5/5500000*100', 'type': 'formula', 'fmt': 'percent'},

    # Row 5: Services
    {'r': 5, 'c': 0, 'v': 'Услуги', 'type': 'label'},
    {'r': 5, 'c': 1, 'v': '300000', 'type': 'input', 'fmt': 'currency'},
    {'r': 5, 'c': 2, 'v': '320000', 'type': 'input', 'fmt': 'currency'},
    {'r': 5, 'c': 3, 'v': '310000', 'type': 'input', 'fmt': 'currency'},
    {'r': 5, 'c': 4, 'v': '340000', 'type': 'input', 'fmt': 'currency'},
    {'r': 5, 'c': 5, 'v': '360000', 'type': 'input', 'fmt': 'currency'},
    {'r': 5, 'c': 6, 'v': '380000', 'type': 'input', 'fmt': 'currency'},
    {'r': 5, 'c': 7, 'v': '=SUM(B6:G6)', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 5, 'c': 8, 'v': '=H6/2000000*100', 'type': 'formula', 'fmt': 'percent'},

    # Row 6: Total Revenue
    {'r': 6, 'c': 0, 'v': 'ИТОГО ДОХОДЫ', 'type': 'label', 'bold': True, 'bg': '#c8e6c9'},
    {'r': 6, 'c': 1, 'v': '=SUM(B4:B6)', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 6, 'c': 2, 'v': '=SUM(C4:C6)', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 6, 'c': 3, 'v': '=SUM(D4:D6)', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 6, 'c': 4, 'v': '=SUM(E4:E6)', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 6, 'c': 5, 'v': '=SUM(F4:F6)', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 6, 'c': 6, 'v': '=SUM(G4:G6)', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 6, 'c': 7, 'v': '=SUM(H4:H6)', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 6, 'c': 8, 'v': '=H7/16500000*100', 'type': 'formula', 'fmt': 'percent', 'bold': True},

    # Row 7: blank separator
    {'r': 7, 'c': 0, 'v': '', 'type': 'label'},

    # Row 8: Section header - Expenses
    {'r': 8, 'c': 0, 'v': '─── РАСХОДЫ ───', 'type': 'label', 'bold': True, 'bg': '#fce4ec'},

    # Row 9: Payroll
    {'r': 9, 'c': 0, 'v': 'Фонд оплаты труда', 'type': 'label'},
    {'r': 9, 'c': 1, 'v': '600000', 'type': 'input', 'fmt': 'currency'},
    {'r': 9, 'c': 2, 'v': '600000', 'type': 'input', 'fmt': 'currency'},
    {'r': 9, 'c': 3, 'v': '620000', 'type': 'input', 'fmt': 'currency'},
    {'r': 9, 'c': 4, 'v': '620000', 'type': 'input', 'fmt': 'currency'},
    {'r': 9, 'c': 5, 'v': '640000', 'type': 'input', 'fmt': 'currency'},
    {'r': 9, 'c': 6, 'v': '640000', 'type': 'input', 'fmt': 'currency'},
    {'r': 9, 'c': 7, 'v': '=SUM(B10:G10)', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 9, 'c': 8, 'v': '=H10/B7*100', 'type': 'formula', 'fmt': 'percent'},

    # Row 10: Rent
    {'r': 10, 'c': 0, 'v': 'Аренда офиса', 'type': 'label'},
    {'r': 10, 'c': 1, 'v': '150000', 'type': 'input', 'fmt': 'currency'},
    {'r': 10, 'c': 2, 'v': '150000', 'type': 'input', 'fmt': 'currency'},
    {'r': 10, 'c': 3, 'v': '150000', 'type': 'input', 'fmt': 'currency'},
    {'r': 10, 'c': 4, 'v': '150000', 'type': 'input', 'fmt': 'currency'},
    {'r': 10, 'c': 5, 'v': '150000', 'type': 'input', 'fmt': 'currency'},
    {'r': 10, 'c': 6, 'v': '150000', 'type': 'input', 'fmt': 'currency'},
    {'r': 10, 'c': 7, 'v': '=SUM(B11:G11)', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 10, 'c': 8, 'v': '=H11/B7*100', 'type': 'formula', 'fmt': 'percent'},

    # Row 11: Marketing
    {'r': 11, 'c': 0, 'v': 'Маркетинг и реклама', 'type': 'label'},
    {'r': 11, 'c': 1, 'v': '200000', 'type': 'input', 'fmt': 'currency'},
    {'r': 11, 'c': 2, 'v': '220000', 'type': 'input', 'fmt': 'currency'},
    {'r': 11, 'c': 3, 'v': '250000', 'type': 'input', 'fmt': 'currency'},
    {'r': 11, 'c': 4, 'v': '230000', 'type': 'input', 'fmt': 'currency'},
    {'r': 11, 'c': 5, 'v': '260000', 'type': 'input', 'fmt': 'currency'},
    {'r': 11, 'c': 6, 'v': '280000', 'type': 'input', 'fmt': 'currency'},
    {'r': 11, 'c': 7, 'v': '=SUM(B12:G12)', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 11, 'c': 8, 'v': '=H12/B7*100', 'type': 'formula', 'fmt': 'percent'},

    # Row 12: Other expenses
    {'r': 12, 'c': 0, 'v': 'Прочие расходы', 'type': 'label'},
    {'r': 12, 'c': 1, 'v': '80000', 'type': 'input', 'fmt': 'currency'},
    {'r': 12, 'c': 2, 'v': '75000', 'type': 'input', 'fmt': 'currency'},
    {'r': 12, 'c': 3, 'v': '90000', 'type': 'input', 'fmt': 'currency'},
    {'r': 12, 'c': 4, 'v': '85000', 'type': 'input', 'fmt': 'currency'},
    {'r': 12, 'c': 5, 'v': '95000', 'type': 'input', 'fmt': 'currency'},
    {'r': 12, 'c': 6, 'v': '100000', 'type': 'input', 'fmt': 'currency'},
    {'r': 12, 'c': 7, 'v': '=SUM(B13:G13)', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 12, 'c': 8, 'v': '=H13/B7*100', 'type': 'formula', 'fmt': 'percent'},

    # Row 13: Total Expenses
    {'r': 13, 'c': 0, 'v': 'ИТОГО РАСХОДЫ', 'type': 'label', 'bold': True, 'bg': '#ffcdd2'},
    {'r': 13, 'c': 1, 'v': '=SUM(B10:B13)', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 13, 'c': 2, 'v': '=SUM(C10:C13)', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 13, 'c': 3, 'v': '=SUM(D10:D13)', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 13, 'c': 4, 'v': '=SUM(E10:E13)', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 13, 'c': 5, 'v': '=SUM(F10:F13)', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 13, 'c': 6, 'v': '=SUM(G10:G13)', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 13, 'c': 7, 'v': '=SUM(H10:H13)', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 13, 'c': 8, 'v': '=H14/H7*100', 'type': 'formula', 'fmt': 'percent', 'bold': True},

    # Row 14: blank
    {'r': 14, 'c': 0, 'v': '', 'type': 'label'},

    # Row 15: Gross Profit
    {'r': 15, 'c': 0, 'v': 'ВАЛОВАЯ ПРИБЫЛЬ', 'type': 'label', 'bold': True, 'bg': '#e3f2fd'},
    {'r': 15, 'c': 1, 'v': '=B7-B14', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 15, 'c': 2, 'v': '=C7-C14', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 15, 'c': 3, 'v': '=D7-D14', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 15, 'c': 4, 'v': '=E7-E14', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 15, 'c': 5, 'v': '=F7-F14', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 15, 'c': 6, 'v': '=G7-G14', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 15, 'c': 7, 'v': '=H7-H14', 'type': 'formula', 'fmt': 'currency', 'bold': True},
    {'r': 15, 'c': 8, 'v': '=H16/H7*100', 'type': 'formula', 'fmt': 'percent', 'bold': True},

    # Row 16: Margin %
    {'r': 16, 'c': 0, 'v': 'Рентабельность, %', 'type': 'label'},
    {'r': 16, 'c': 1, 'v': '=B16/B7*100', 'type': 'formula', 'fmt': 'percent'},
    {'r': 16, 'c': 2, 'v': '=C16/C7*100', 'type': 'formula', 'fmt': 'percent'},
    {'r': 16, 'c': 3, 'v': '=D16/D7*100', 'type': 'formula', 'fmt': 'percent'},
    {'r': 16, 'c': 4, 'v': '=E16/E7*100', 'type': 'formula', 'fmt': 'percent'},
    {'r': 16, 'c': 5, 'v': '=F16/F7*100', 'type': 'formula', 'fmt': 'percent'},
    {'r': 16, 'c': 6, 'v': '=G16/G7*100', 'type': 'formula', 'fmt': 'percent'},
    {'r': 16, 'c': 7, 'v': '=AVERAGE(B17:G17)', 'type': 'formula', 'fmt': 'percent'},
    {'r': 16, 'c': 8, 'v': '', 'type': 'label'},
]


def seed_demo_data():
    """Create demo financial table in the database."""
    sheet, _ = Sheet.objects.get_or_create(name='Демо Бюджет')
    sheet.cells.all().delete()
    sheet.dependencies.all().delete()

    from .formula_engine import make_eval_context

    cells_to_create = []
    formula_cells = []

    for item in DEMO_DATA:
        r, c = item['r'], item['c']
        raw_v = item['v']
        cell_type = item.get('type', 'input')
        fmt = item.get('fmt', '')
        bold = item.get('bold', False)
        bg = item.get('bg', '')
        col_span = item.get('col_span', 1)
        row_span = item.get('row_span', 1)

        formula = None
        py_formula = None
        value = raw_v

        if isinstance(raw_v, str) and raw_v.startswith('='):
            formula = raw_v
            py_formula = formula_to_python(raw_v[1:])
            cell_type = 'formula'
            value = '0'
            formula_cells.append((r, c, raw_v))

        dec = 2 if fmt in ('currency', 'number') else (1 if fmt == 'percent' else 0)
        is_editable = cell_type == 'input'

        cells_to_create.append(Cell(
            sheet=sheet,
            row=r,
            col=c,
            value=value,
            raw_value=raw_v,
            formula=formula,
            python_formula=py_formula,
            cell_type=cell_type,
            is_editable=is_editable,
            format_type=fmt,
            decimal_places=dec,
            col_span=col_span,
            row_span=row_span,
            bold=bold,
            bg_color=bg,
        ))

    Cell.objects.bulk_create(cells_to_create)

    # Build dependencies
    deps = []
    for (row_idx, col_idx, formula) in formula_cells:
        refs = extract_cell_refs(formula)
        for (ref_r, ref_c) in refs:
            deps.append(CellDependency(
                sheet=sheet,
                source_row=ref_r,
                source_col=ref_c,
                target_row=row_idx,
                target_col=col_idx,
            ))
    if deps:
        CellDependency.objects.bulk_create(deps, ignore_conflicts=True)

    # Full initial recalc
    from .formula_engine import recalculate_sheet
    changed, cell_values = recalculate_sheet(sheet)
    if changed:
        all_cells = list(sheet.cells.filter(cell_type='formula'))
        cell_map = {(c.row, c.col): c for c in all_cells}
        to_save = []
        for (r, c), val in changed.items():
            cell = cell_map.get((r, c))
            if cell:
                cell.value = val
                to_save.append(cell)
        if to_save:
            Cell.objects.bulk_update(to_save, ['value'])

    return sheet
