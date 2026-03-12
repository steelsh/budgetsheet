from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('spreadsheet', '0001_initial'),
        ('auth', '0012_alter_user_first_name_max_length'),
    ]

    operations = [
        # Add comment to Cell
        migrations.AddField(
            model_name='cell',
            name='comment',
            field=models.TextField(blank=True, null=True),
        ),
        # Add owner to Sheet
        migrations.AddField(
            model_name='sheet',
            name='owner',
            field=models.ForeignKey(
                blank=True, null=True,
                on_delete=django.db.models.deletion.SET_NULL,
                related_name='sheets',
                to='auth.user',
            ),
        ),
        # Add is_template to Sheet
        migrations.AddField(
            model_name='sheet',
            name='is_template',
            field=models.BooleanField(default=False),
        ),
        # Create SheetSnapshot
        migrations.CreateModel(
            name='SheetSnapshot',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(max_length=255)),
                ('note', models.TextField(blank=True, default='')),
                ('data', models.JSONField()),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('created_by', models.CharField(blank=True, default='anonymous', max_length=255)),
                ('sheet', models.ForeignKey(
                    on_delete=django.db.models.deletion.CASCADE,
                    related_name='snapshots',
                    to='spreadsheet.sheet',
                )),
            ],
            options={
                'verbose_name': 'Снимок',
                'verbose_name_plural': 'Снимки',
                'ordering': ['-created_at'],
            },
        ),
    ]
