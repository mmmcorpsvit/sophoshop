from openpyxl import load_workbook
import xmlrpc.client as xmlrpclib
import environ
import base64

# import os
# from os import rename
# import tarfile
# import zipfile
# import zlib
# import shutil
# import tempfile


def convert_image(fn):
    with open(fn, "rb") as image_file:
        return base64.b64encode(image_file.read())


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


# c = Impxls()
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

    def unlink_item(self):
        l2 = [10, 11, 20, 26, 21]

        # check if the deleted record is still in the database
        lst3 = self._models_objects.execute_kw(
            self._db, self._uid, self._password,
            'product.template',
            'search',
            [
                [
                    ['id', 'not in', l2]
                ]
            ]
        )

        for e2 in lst3:
            lst2 = self._models_objects.execute_kw(
                self._db, self._uid, self._password,
                'product.template',
                'unlink',
                [e2])

        out('\n==unlink items==')
        out('count: %i' % len(lst3))
        pass

    def unlink_attributes(self):
        l2 = [1, 2, 3]

        attribute_id = self._models_objects.execute_kw(
            self._db, self._uid, self._password,
            'product.attribute',  # model (just see param model in admin side URL)
            'search_read',  # operation
            [[
                # conditions
                # ['parent_id', 'in', [False, ]],  # one item
                ['id', 'not in', l2],  # item in set
              ],

                ['id']  # fields list
             ]
            )

        for e2 in attribute_id:
            lst2 = self._models_objects.execute_kw(
                self._db, self._uid, self._password,
                'product.attribute',
                'unlink',
                [e2])

        pass

    def create_attribute(self, sname, svalue):
        # search attribute
        attribute_id = self._models_objects.execute_kw(
            self._db, self._uid, self._password,
            'product.attribute',  # model (just see param model in admin side URL)
            'search_read',  # operation
            [[
                # conditions
                # ['parent_id', 'in', [False, ]],  # one item
                ['name', 'in', [sname]],  # item in set
              ],

                ['id']  # fields list
             ]
            )

        # create attribute if not exists
        if not attribute_id:
            attribute_id = self._models_objects.execute_kw(
                self._db, self._uid, self._password,
                'product.attribute',
                'create',
                [{
                    'name': sname,
                }]
            )

        id = attribute_id[0]['id']

        # search attribute value
        value_attribute_id = self._models_objects.execute_kw(
            self._db, self._uid, self._password,
            'product.attribute.value',  # model (just see param model in admin side URL)
            'search_read',  # operation
            [[
                # conditions
                # ['parent_id', 'in', [False, ]],  # one item
                ['attribute_id', 'in', [id]],
                ['name', 'in', [svalue]],  # item in set
              ],

                # ['id']  # fields list
             ]
            )

        # create attribute if not exists
        if not value_attribute_id:
            value_attribute_id = self._models_objects.execute_kw(
                self._db, self._uid, self._password,
                'product.attribute.value',
                'create',
                [{
                    'attribute_id': id,
                    'name': svalue,
                    # 'html_color': False,
                }]
            )

        return value_attribute_id

    def create_item(self, item):
        """
        Create in from lst
        :param item: list of dict(name, cat_name, price, desc)
        :return:
        """
        # for e in lst:
        # i += 1
        categ_id = int(self._cats[item['cat_name']])
        sname = item['name']

        product_id = self._models_objects.execute_kw(
                self._db, self._uid, self._password,
                'product.template', 'create',
                [{
                   'name': sname,
                   'price': item['price'],
                   'categ_id': 6,  # All / Можна продавати / Physical
                   # 'default_code': '1111',
                   'public_categ_ids': [[6, 0, [categ_id]]],
                   # 'description_sale': 'super_puper_long',
                   'website_description': item['desc'],
                   'website_published': True,
                   'image': item['image'],
                }]
                )

        # add attributes (pnly if have his)
        if hasattr(item, 'attributes'):
            attr = item['attributes']
            for e in attr:
                self.create_attribute(e, attr[e])


        # out('[id: %i] [%s] ' % (product_id, sname,))

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
     'image': '/9j/4AAQSkZJRgABAQEBLAEsAAD/2wBDAA0JCgsKCA0LCgsODg0PEyAVExISEyccHhcgLikxMC4pLSwzOko+MzZGNywtQFdBRkxOUlNSMj5aYVpQYEpRUk//2wBDAQ4ODhMREyYVFSZPNS01T09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0//wAARCABLAGQDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDaSPHWpVIXrzTRijFd7Z8+oWJ1cdl/OniQ+uPpVdTUygnpUNmsVcfml3GmfMP4TxSDJ6Url8jJlfHenh81CqnvVlI/WpbNIxYqgtUyoF60zO3gClBYjIIpXLUSXfSb6iyR1pM0h3ZN5lFRUUCuzEVvWn5psILDkDNSiMg8g03ViZrDT7DVzUqMVOaarITjcM1OqZqHVTNoYeSZJA43fNzmpzDGRno3aoEUA4NWFAx3rGU0nc7YQbjZobs24LCp0KhQzHNM6nlc0hG0cdDUyqX0LhRSdyycMuDg8cVCML0HNLGxAwelKQM8URlbQKkb6gy7xxxTDEV7VMoIHNBzTVSxMqKav1K+KKl4oqvaoy9gzCmi2j7pB9arTa7p9gPJvrhYpOv3WJP5CtYgM3zdKqXdrDKpjlijkQ/wuuRXnKok9T1XC60KA1fSbzDW97EW9G+Qn88Vc/tC0tQv2i6iTJ4BbOax7vwrp8/zW4a2c/3TlfyP9K5670TUNPG6aEtGP44+VH19PxrZTjLZmXs2t0em09TxmuAstf1WAAfaTIo7SKG/XrVi41K+1FgsrnDHAjQYH5d6lytuUqbOqudbsbckGXzW7iPn9elFhrdvf3At44ZgxBOSBgfXms2w8Mu2HvpNg/uLyfxNdDa2cFomy3iWMd8dT9TS5waSJgvpS4xQBinUc5FhM+lHWlxSYqlIloTaKKdj2oo5gMgHBzXOa/rs0EEiQWt/BIp+WdoAUP4muhBpSFdSjgMrDBBHBFc8ZK92jfU4mx8YXONlxbpO/RSnykn6c1ZeDxDrbDzIHih7K37tR+B5P61vQ6Dp0WqR38MIikjXARAFTPTOPWtpTmqlOK1SBNnITeGpbPT5bme5QugB2KDjqO//ANarvhnT5jcx3TQAw84Zj39QK6R4o5kMUy7kbqD3qeNVRAqKFUDAAGAKlS5huTSsSYFKBQKeBVJXMWxu2jbUm2jbVcpNyPFLTsUhpgJRRiigdjn1JIxin5xUaDGMVJ3NcEajud0oJDlYVLHJ2/WoVqRapyuZ2sW0YGp1Oaqx1YjqoszkWFqUCokqZa6oHPIXFBFLRW3KQRkU0ipDTDUuJaGUUUVnYo//2Q==',

     'attributes': {
        'цвет': 'красный',
        'форма': 'квадрат',
        'материал': 'бук',
        }
     },

    {'name': 'бозен2',
     'price': 66,
     'cat_name': 'диван',
     'desc': 'super duper divan2',
     # 'cat_name': 'диван',
     'image': '',
     },


]

im = ImportToOdd('http://localhost:8069',
                 'shop2',
                 env('user'),
                 env('pwd'))
im.connect()

# im.test()
# im.unlink_attributes()
# im.create_attribute('форма', 'квадрат')


cats = im.create_categories(list_1)


# im.unlink_item()

out('\n==products==')
i = 0
for e in list_2:
    i += 1
    out('[%i/%i] [%s]' % (i, len(list_2), e['name'],))
    im.create_item(e)



