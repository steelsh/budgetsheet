from django.contrib import admin
from .models import Sheet, Cell, CellDependency, ChangeHistory


@admin.register(Sheet)
class SheetAdmin(admin.ModelAdmin):
    list_display = ['name', 'created_at', 'updated_at']


@admin.register(Cell)
class CellAdmin(admin.ModelAdmin):
    list_display = ['sheet', 'excel_ref', 'cell_type', 'value', 'is_editable', 'formula']
    list_filter = ['sheet', 'cell_type', 'is_editable']
    search_fields = ['value', 'formula']


@admin.register(CellDependency)
class DependencyAdmin(admin.ModelAdmin):
    list_display = ['sheet', 'source_row', 'source_col', 'target_row', 'target_col']


@admin.register(ChangeHistory)
class HistoryAdmin(admin.ModelAdmin):
    list_display = ['cell', 'old_value', 'new_value', 'changed_at', 'changed_by']
    list_filter = ['changed_at']
