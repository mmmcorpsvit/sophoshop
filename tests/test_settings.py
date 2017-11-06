from django.test import TestCase
from django.utils.translation import ugettext_lazy as _

# DATABASES['default'] = DB.DATABASES['default']


class TestSettings(TestCase):
    def test_ukraine_language(self):
        print(_("Body Text"))
        self.assertTrue(_("Body Text") == "Текст Тіла", 'Ukraine language fix error')
        pass
