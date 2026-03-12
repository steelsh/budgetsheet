from django.db import models
from django.contrib.auth.models import User


class Sheet(models.Model):
    name = models.CharField(max_length=255, default='Бюджет')
    owner = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='sheets')
    is_template = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return self.name

    class Meta:
        verbose_name = 'Лист'
        verbose_name_plural = 'Листы'
        ordering = ['-updated_at']


class Cell(models.Model):
    CELL_TYPES = [
        ('input', 'Ввод'),
        ('formula', 'Формула'),
        ('label', 'Метка'),
        ('header', 'Заголовок'),
    ]
    sheet = models.ForeignKey(Sheet, on_delete=models.CASCADE, related_name='cells')
    row = models.IntegerField()
    col = models.IntegerField()
    value = models.TextField(blank=True, null=True)
    raw_value = models.TextField(blank=True, null=True)
    formula = models.TextField(blank=True, null=True)
    python_formula = models.TextField(blank=True, null=True)
    cell_type = models.CharField(max_length=10, choices=CELL_TYPES, default='input')
    is_editable = models.BooleanField(default=False)
    format_type = models.CharField(max_length=50, blank=True, default='')
    decimal_places = models.IntegerField(default=2)
    row_span = models.IntegerField(default=1)
    col_span = models.IntegerField(default=1)
    css_class = models.CharField(max_length=255, blank=True, default='')
    bold = models.BooleanField(default=False)
    italic = models.BooleanField(default=False)
    bg_color = models.CharField(max_length=20, blank=True, default='')
    text_color = models.CharField(max_length=20, blank=True, default='')
    number_format = models.CharField(max_length=100, blank=True, default='')
    comment = models.TextField(blank=True, null=True)

    class Meta:
        unique_together = ('sheet', 'row', 'col')
        ordering = ['row', 'col']
        verbose_name = 'Ячейка'
        verbose_name_plural = 'Ячейки'

    def __str__(self):
        return f"{self.sheet.name}[{self.row},{self.col}] = {self.value}"

    @property
    def col_letter(self):
        result = ''
        n = self.col
        while True:
            result = chr(65 + n % 26) + result
            n = n // 26 - 1
            if n < 0:
                break
        return result

    @property
    def excel_ref(self):
        return f"{self.col_letter}{self.row + 1}"


class CellDependency(models.Model):
    sheet = models.ForeignKey(Sheet, on_delete=models.CASCADE, related_name='dependencies')
    source_row = models.IntegerField()
    source_col = models.IntegerField()
    target_row = models.IntegerField()
    target_col = models.IntegerField()

    class Meta:
        unique_together = ('sheet', 'source_row', 'source_col', 'target_row', 'target_col')
        verbose_name = 'Зависимость'
        verbose_name_plural = 'Зависимости'


class ChangeHistory(models.Model):
    cell = models.ForeignKey(Cell, on_delete=models.CASCADE, related_name='history')
    old_value = models.TextField(blank=True, null=True)
    new_value = models.TextField(blank=True, null=True)
    changed_at = models.DateTimeField(auto_now_add=True)
    changed_by = models.CharField(max_length=255, blank=True, default='anonymous')

    class Meta:
        ordering = ['-changed_at']
        verbose_name = 'История изменений'
        verbose_name_plural = 'История изменений'


class SheetSnapshot(models.Model):
    """Full snapshot of a sheet for version history / rollback"""
    sheet = models.ForeignKey(Sheet, on_delete=models.CASCADE, related_name='snapshots')
    name = models.CharField(max_length=255)
    note = models.TextField(blank=True, default='')
    data = models.JSONField()  # serialized cells
    created_at = models.DateTimeField(auto_now_add=True)
    created_by = models.CharField(max_length=255, blank=True, default='anonymous')

    class Meta:
        ordering = ['-created_at']
        verbose_name = 'Снимок'
        verbose_name_plural = 'Снимки'

    def __str__(self):
        return f"{self.sheet.name} @ {self.created_at:%d.%m.%Y %H:%M}"
