from openpyxl import load_workbook
import xmlrpc.client as xmlrpclib
import environ

# import os
# from os import rename
# import tarfile
# import zipfile
# import zlib
# import shutil
# import tempfile

env = environ.Env()
out = print

# region 'conditions list const'
CON_IGNORE_SNAME_STARTS = ['Тканини для асортименту "ТИСА-МЕБЛІ"',
                           ]

CON_IGNORE_ATTRS = ['Гарантийный срок',
                    'Ручки для переноса',
                    ]


# endregion


class Impxls(object):
    _flush = False
    _add_images = False
    _rebuild_index = False
    _csv1 = None
    _csv2 = None

    def __init__(self, flush=False, add_images=False, rebuild_index=False):
        self._flush = flush
        self._add_images = add_images
        self._rebuild_index = rebuild_index
        self._csv1 = open('attr1.csv', 'w')
        self._csv2 = open('attr2.csv', 'w')

    def handle(self, fn):
        wb2 = load_workbook(fn, read_only=True)
        # print('Spread sheats names: %s' % wb2.get_sheet_names())
        wb = wb2.worksheets[0]

        index = 0
        cats = dict()
        attr_dict = dict()

        extra_attr_dict = dict()  # name, [type, value]

        total_count = wb.max_row

        for row in wb.rows:
            extra_attr_dict.clear()
            # region 'work'
            index += 1

            if index == 1:
                continue

            # if index < 103:
            #     continue

            # if index > 50:
            #     break

            v = row[5].value
            price = int(v) if not (v is None) else 0

            v = row[3].value
            description = str(v) if not (v is None) else ''

            cat_original = str(row[15].value)

            cat = cat_original \
                .replace('Матраци Sleep&Fly', 'Матраци') \
                .replace('Матраци Evolution', 'Матраци') \
                .replace('Матраци Sleep&fly Organic', 'Матраци') \
                .replace('Матраци Take&Go Bamboо', 'Матраци') \
                .replace('Take&Go', 'Матраци') \
                .replace('Матраци Sleep&fly uno', 'Матраци') \
                .replace('Матраци Doctor Health', 'Матраци') \
                .replace('Дитячі матраци Herbalis KIDS', 'Матраци') \
                .replace('Матраци на дивани', 'Футони і топери') \
                .replace('Наматрацники', 'Наматрацники і підматрацники') \
                .replace("Дерев'яні ліжка", 'Ліжка') \
                .replace("Дитячі ліжка", 'Ліжка') \
                .replace('Столи', 'Столи гостьові') \
                .replace('Столи гостьові-трансформери', 'Столи журнальні')\
                .replace('Стільці', 'Стільці та табурети') \
                .replace('Дитячі дивани', 'Дивани') \
                .replace('Кутові дивани', 'Дивани') \
                .replace('Прямі дивани', 'Дивани') \
                .strip()
            cats[cat] = None

            brand = str(row[24].value) \
                .replace('Скиф', 'Скіф') \
                .replace('Тиса мебель', 'Тиса меблі') \
                .replace('Елисеевская мебель', 'Єлисеївські меблі') \
                .replace('Микс мебель', 'Мікс меблі') \
                .replace('Мелитополь мебель', 'Мелітополь меблі')

            country_manufactur = str(row[26].value) \
                .replace('Украина', 'Україна')

            some_list = list(row[30:])

            # add garanty 18 month attribute
            if cat in [
                'Матраци',
                'Дивани',
                'Подушки',
                'Наматрацники і підматрацники',
                'Футони і топери',
                 ]:
                # extra_attr_dict['Гарантыя'] = '18'  # просто добавляем в словарь без типов данных
                extra_attr_dict['Гарантія'] = ['int', '18']

            sname = str(row[1].value)

            # ignore some "with start" sname products
            for e in CON_IGNORE_SNAME_STARTS:
                if sname.startswith(e):
                    sname = ''
            if sname == '':
                continue

            # if sname.startswith('Тканини для асортименту "ТИСА-МЕБЛІ"'):
            #    continue

            sname = sname \
                .strip()\
                .replace('"', '')\
                .replace('  ', ' ')\
                .replace('*', 'x')
            out('[%i/%i] %s' % (index, total_count, sname))

            af = []
            # get list of tuples(name_product, name_attr, name_counter, value)
            # use step by 1

            # region 'convert work'
            attrs_list0 = [[sname,
                           str(val.value).strip(),
                           str(some_list[idx+1].value).strip(),
                           str(some_list[idx+2].value).strip()
                            .replace('.0', '')
                            .replace('*', 'x')
                            ]
                           for idx, val in enumerate(some_list)
                           if idx % 3 == 0
                           and val.value is not None]

            # skip ignored attributes
            attrs_list1 = [e for e in attrs_list0
                           if e[1] not in CON_IGNORE_ATTRS
                           and not(e[1] == 'Цвет' and e[3] == 'Разные цвета')
                           and not(e[1] == 'Розмір' and e[3] == '7000')
                           and not(e[1] == 'Состояние' and e[3] == 'Новое')
                           and not(e[1] == 'Тип' and e[3] == 'Для сна')
                           and not(e[1] == 'Цвет' and e[3] == 'Белый')
                           and not(e[1] == 'Цвет' and e[3] == 'Разные цвета')
                           and not(e[1] == 'Цвет обивки' and e[3] == 'Разные цвета')
                           and not(e[1] == 'Тип крепления к матрасу' and e[3] == 'четыре резинки по углам')
                           ]

            # attrs_list2 = []

            # mm to sm
            for idx, e in enumerate(attrs_list1):
                if cat in ['Столи гостьові',
                           'Стільці та табурети', ] \
                        and e[1] in ['Глубина столика',
                                     'Длина столика',
                                     'Максимальная длина столешницы раскладного столика',
                                     'Минимальная длина столешницы раскладного столика',
                                     'Длина стола',
                                     'Высота',
                                     'Длина стола в раздвинутом (разложенном) состоянии',
                                     'Длина стола в сдвинутом (сложенном) состоянии',
                                     'Ширина',
                                     'Глубина',
                                     'Ширина стола',
                                     ]:
                    e[3] = str(int(e[3])/10).replace('.0', '')
                    e[2] = 'см'

                if e[2] in ['см', 'кг', 'шт.']:
                    e[2] = 'int'

                if e[3] in ['да', 'нет']:
                    e[2] = 'bool'

                    e[3] = e[3].replace('да', '1').replace('нет', '0')

                # else string
                # if e[2] == '':
                #    e[2] = 'str'

                # attrs_list2.append(result)

            attrs_list = attrs_list1

            # add extra data
            for e in extra_attr_dict:
                attrs_list.append([sname, e, extra_attr_dict[e][0], extra_attr_dict[e][1]])

            # detect types of data

            # endregion

            for e in attrs_list:
                s = '%s\t%s\t%s\t%s\n' % (e[0], e[1], e[2], e[3])
                self._csv1.writelines(s)

                # just for debug
                attr_dict[e[1]] = ''
                # self._csv += e[0]+'\t'+e[1]+'\t'+e[2]+'\n'

            # print('[%i/%i] [%s] %s' % (index - 1, wb.max_row, cat, row[1].value,))

            for e in attrs_list:
                pass
                # print(e[0])

        self._csv1.close()
        self._csv2.close()
        wb2.close()

        out('\n\n==Cats==')
        for e in cats:
            out(e)

        out('\n\n==Attr_dict==')
        # attr_dict.
        for e in sorted(attr_dict):
            out(e)


c = Impxls()
# c.handle('export-products.xlsx')


# http://www.odoo.com/documentation/10.0/api_integration.html
class ImportToOdd:
    _srv, _db, _username, _password = '', '', '', ''
    _uid = None
    _models_objects = None
    _common_objects = None
    _cats = dict()

    def __init__(self, _srv, _db, _username, _password):
        self._srv, self._db, self._username, self._password = _srv, _db, _username, _password
        # self._cats = cats

    def connect(self):
        self._cats.clear()
        rpc = '%s/xmlrpc/2/' % self._srv

        self._common_objects = xmlrpclib.ServerProxy('%scommon' % rpc)
        cv = self._common_objects.version()
        out('\n==server info==')
        out(cv)

        self._uid = self._common_objects.authenticate(self._db, self._username, self._password, {})
        if not self._uid:
            Exception('Authorization error')
        else:
            out('Authorization succes, uid: %s' % self._uid)

        self._models_objects = xmlrpclib.ServerProxy('%sobject' % rpc)

    def create_categories(self, lst):
        """
        ReCreate Categories list in root
        :param lst: list of names categories
        :return: dict(name, id)
        """
        result = dict()
        # get cats list to delete
        categories = self._models_objects.execute_kw(
            self._db, self._uid, self._password,
            'product.public.category',  # model (just see param model in admin side URL)
            'search_read',  # operation
            [[
                # conditions
                ['parent_id', 'in', [False, ]],  # one item
                ['name', 'in', lst],  # item in set
              ],

                ['id']  # fields list
             ]
            )

        # delete
        ids = []
        for e in categories:
            ids.append(e['id'])

        self._models_objects.execute_kw(
            self._db, self._uid, self._password,
            'product.public.category',

            # operation
            'unlink',  # delete
            [ids]  # from list of id's
        )

        # create
        for e in list_1:
            result[e] = self._models_objects.execute_kw(
                self._db, self._uid, self._password,
                'product.public.category',

                # operation
                'create',  # delete
                [{'name': e}]  # from list of id's
            )

        out('\n==cats==')
        for e in result:
            out('%s: %i' % (e, result[e]))

        self._cats = result
        return result

    def create_item(self, lst):
        """
        Create in from lst
        :param lst: list of dict(name, cat_name, price, desc)
        :return:
        """
        i = 0
        cnt = len(lst)

        # create
        out('\n==products==')
        for e in lst:
            i += 1
            categ_id = int(self._cats[e['cat_name']])
            sname = e['name']

            product_id = self._models_objects.execute_kw(
                    self._db, self._uid, self._password,
                    'product.template', 'create',
                    [{
                       'name': sname,
                       'price': e['price'],
                       'public_categ_id': categ_id,
                       # 'description_sale': 'super_puper_long',
                       'website_description': e['desc'],
                       'website_published': True,
                       # 'image': None,
                    }]
                    )

            out('[%i/%i] [id: %i] [%s] ' % (i, cnt, product_id, sname,))

    def test(self):
        # just see in Firefox + F12 debug, template + method + params

        product = self._models_objects.execute_kw(
            self._db, self._uid, self._password,
            'product.template', 'read',
            [
                # 3,
                10,
                # product_id,
            ]
            )[0]

        pass


# test
list_1 = ['диван', 'Network']
list_2 = [
    {'name': 'бозен',
     'price': 99,
     'cat_name': 'диван',
     'desc': 'super duper divan',
     # 'cat_name': 'диван',
     },

    {'name': 'бозен2',
     'price': 66,
     'cat_name': 'диван',
     'desc': 'super duper divan2',
     # 'cat_name': 'диван',
     },


]

im = ImportToOdd('http://localhost:8069',
                 'shop',
                 env('user'),
                 env('pwd'))
im.connect()
cats = im.create_categories(list_1)

im.create_item(list_2)
# im.test()
