from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    initial = True

    dependencies = [
    ]

    operations = [
        migrations.CreateModel(
            name='Sheet',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(default='Бюджет', max_length=255)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('updated_at', models.DateTimeField(auto_now=True)),
            ],
            options={
                'verbose_name': 'Лист',
                'verbose_name_plural': 'Листы',
            },
        ),
        migrations.CreateModel(
            name='Cell',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('row', models.IntegerField()),
                ('col', models.IntegerField()),
                ('value', models.TextField(blank=True, null=True)),
                ('raw_value', models.TextField(blank=True, null=True)),
                ('formula', models.TextField(blank=True, null=True)),
                ('python_formula', models.TextField(blank=True, null=True)),
                ('cell_type', models.CharField(choices=[('input', 'Ввод'), ('formula', 'Формула'), ('label', 'Метка'), ('header', 'Заголовок')], default='input', max_length=10)),
                ('is_editable', models.BooleanField(default=False)),
                ('format_type', models.CharField(blank=True, default='', max_length=50)),
                ('decimal_places', models.IntegerField(default=2)),
                ('row_span', models.IntegerField(default=1)),
                ('col_span', models.IntegerField(default=1)),
                ('css_class', models.CharField(blank=True, default='', max_length=255)),
                ('bold', models.BooleanField(default=False)),
                ('italic', models.BooleanField(default=False)),
                ('bg_color', models.CharField(blank=True, default='', max_length=20)),
                ('text_color', models.CharField(blank=True, default='', max_length=20)),
                ('number_format', models.CharField(blank=True, default='', max_length=100)),
                ('sheet', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='cells', to='spreadsheet.sheet')),
            ],
            options={
                'verbose_name': 'Ячейка',
                'verbose_name_plural': 'Ячейки',
                'ordering': ['row', 'col'],
                'unique_together': {('sheet', 'row', 'col')},
            },
        ),
        migrations.CreateModel(
            name='CellDependency',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('source_row', models.IntegerField()),
                ('source_col', models.IntegerField()),
                ('target_row', models.IntegerField()),
                ('target_col', models.IntegerField()),
                ('sheet', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='dependencies', to='spreadsheet.sheet')),
            ],
            options={
                'verbose_name': 'Зависимость',
                'verbose_name_plural': 'Зависимости',
                'unique_together': {('sheet', 'source_row', 'source_col', 'target_row', 'target_col')},
            },
        ),
        migrations.CreateModel(
            name='ChangeHistory',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('old_value', models.TextField(blank=True, null=True)),
                ('new_value', models.TextField(blank=True, null=True)),
                ('changed_at', models.DateTimeField(auto_now_add=True)),
                ('changed_by', models.CharField(blank=True, default='anonymous', max_length=255)),
                ('cell', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='history', to='spreadsheet.cell')),
            ],
            options={
                'verbose_name': 'История изменений',
                'verbose_name_plural': 'История изменений',
                'ordering': ['-changed_at'],
            },
        ),
    ]
