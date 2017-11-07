from django.test import TestCase
# from sophoshop.settings import *
from django.utils.translation import ugettext_lazy as _


class TestSettings(TestCase):
    def setUp(self):
        pass

    def test_ukraine_language(self):
        self.assertTrue(_("Body Text") == "Текст Тіла", 'Ukraine language fix error')
