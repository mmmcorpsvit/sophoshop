from django.core.management.base import BaseCommand, CommandError
# from polls.models import Question as Poll


class Command(BaseCommand):
    help = 'Import data from csv/xls to database'

    # def add_arguments(self, parser):
    #    parser.add_argument('poll_id', nargs='+', type=int)

    def handle(self, *args, **options):
        self.stdout.write(self.style.SUCCESS('Successfully closed poll "%s"' % 0))
