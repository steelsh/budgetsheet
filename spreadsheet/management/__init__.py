from django.core.management.base import BaseCommand


class Command(BaseCommand):
    help = 'Seed demo financial budget data'

    def handle(self, *args, **options):
        from spreadsheet.demo_data import seed_demo_data
        self.stdout.write('Seeding demo data...')
        sheet = seed_demo_data()
        self.stdout.write(self.style.SUCCESS(
            f'✓ Demo sheet "{sheet.name}" created with {sheet.cells.count()} cells'
        ))
