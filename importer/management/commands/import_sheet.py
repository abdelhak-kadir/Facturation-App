from django.core.management.base import BaseCommand
from importer.utils.sheet_importer import import_livraisons_from_sheet
from importer.utils.sheet_importer2 import import_livraisons_from_sheet2

class Command(BaseCommand):
    help = 'Import data from Google Sheet into Livraison model'

    def handle(self, *args, **kwargs):
        import_livraisons_from_sheet()
        self.stdout.write(self.style.SUCCESS('✅ Livraisons imported successfully!'))
        import_livraisons_from_sheet2()
        self.stdout.write(self.style.SUCCESS('✅ Livraisons prrr imported successfully!'))

