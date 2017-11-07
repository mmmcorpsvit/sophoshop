from django.test import TestCase
# from sophoshop.settings import *
from django.utils.translation import ugettext_lazy as _


class TestSettings(TestCase):
    def setUp(self):
        pass

    def test_ukraine_language(self):
        print(_("Body Text"))
        print('need fix: https://github.com/django-oscar/django-oscar/issues/2465')
        # self.assertTrue(_("Body Text") == "Текст Тіла", 'Ukraine language fix error')
        pass
