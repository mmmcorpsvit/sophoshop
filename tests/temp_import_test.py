# -- encoding: utf8 --

from openpyxl import load_workbook
import xmlrpc.client as xmlrpclib
from environs import Env
import base64

# import datetime
# import os
# from os import rename
# import tarfile
# import zipfile
# import zlib
# import shutil
# import tempfile


def base64_of_file(fn):
    with open(fn, "rb") as image_file:
        return base64.b64encode(image_file.read())


env = Env()
out = print


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
        self.connect()

    def connect(self):
        self._cats.clear()
        rpc = '%s/xmlrpc/2/' % self._srv

        self._common_objects = xmlrpclib.ServerProxy('%scommon' % rpc)
        cv = self._common_objects.version()
        out('\n==server info==')
        out(cv)

        try:
            self._uid = self._common_objects.authenticate(self._db, self._username, self._password, {})
        except:
            pass

        if not self._uid:
            out('Authorization FAILED!')
            exit(-1)

        out('Authorization succes, uid: %s' % self._uid)
        self._models_objects = xmlrpclib.ServerProxy('%sobject' % rpc)

    def create_category(self, cat_name):
        """
        ReCreate Categories list in root (use internal _self.cats for speed)
        :param cat_name: list of names categories
        :return: id
        """
        # already present in list?
        if cat_name in self._cats:
            return self._cats[cat_name]

        # get cats list
        categories = self._models_objects.execute_kw(
            self._db, self._uid, self._password,
            'product.public.category',  # model (just see param model in admin side URL)
            'search_read',  # operation
            [[
                # conditions
                ['parent_id', 'in', [0, ]],  # one item
                ['name', 'in', [cat_name]],  # item in set
            ],
                ['id']  # fields list
            ]
        )

        # categories = [1, 2]
        # category dont exists, create him
        if len(categories) > 0:
            tid = categories[0]['id']
            self._cats[cat_name] = tid
            return tid

        # create
        result = self._models_objects.execute_kw(
            self._db, self._uid, self._password,
            'product.public.category',

            'create',  # delete
            [{'name': cat_name}]  # from list of id's
        )
        out('Create category: %s, id: %i' % (cat_name, result))
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

    def create_attribute(self, attrs_lines, sname, svalue):
        # search attribute
        # svalue = 'test2'
        # region 'scroll'
        attribute_id = self._models_objects.execute_kw(
            self._db, self._uid, self._password,
            'product.attribute',  # model (just see param model in admin side URL)
            'search_read',  # operation
            [[
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
            # attribute_id = attribute_id[0]['id']
        else:
            attribute_id = attribute_id[0]['id']
        # endregion
        # search attribute value
        value_attribute_id = self._models_objects.execute_kw(
            self._db, self._uid, self._password,
            'product.attribute.value',  # model (just see param model in admin side URL)
            'search_read',  # operation
            [[
                # conditions
                # ['parent_id', 'in', [False, ]],  # one item
                ['attribute_id', 'in', [attribute_id]],
                ['name', 'in', [svalue]],  # item in set
            ],
                ['id']  # fields list
            ]
        )

        # create attribute value if not exists
        if not value_attribute_id:
            value_attribute_id = self._models_objects.execute_kw(
                self._db, self._uid, self._password,
                'product.attribute.value',
                'create',
                [{
                    'attribute_id': attribute_id,
                    'name': svalue,
                    # 'html_color': False,
                }]
            )
        else:
            value_attribute_id = value_attribute_id[0]['id']

        attrs_lines += [
            (0, 0,   # what it is?
             {'attribute_id': attribute_id,  # attribute id
              'value_ids': [(4, value_attribute_id), ]},  # [(unknown ???, value_ids)]
             ),
        ]

        return attribute_id, value_attribute_id

    def create_item(self, item):
        """
        Create in from lst
        :param item: list of dict(name, cat_name, price, desc)
        :return:
        """
        # for e in lst:
        # i += 1
        # categ_id = int(self._cats[item['cat_name']])

        # create_categories
        # cat_name = item['cat_name']

        cat_id = self.create_category(item['cat_name'])
        sname = item['name']

        # add attributes (only if have his)
        attrs_lines = []
        try:
            for ekey, evalue in item['attributes'].items():
                self.create_attribute(attrs_lines, ekey, evalue)
        except KeyError:
            # pass
            print('item: %s, dont have attributes!' % sname)

        # ********************
        """
        attrs_lines += [
            (0, 0,   # what it is?
             {'attribute_id': 3,  # attribute id
              'value_ids': [(4, 5), ]},  # [(unknown ???, value_ids)]
             ),

            (0, 0,  # what it is?
             {'attribute_id': 1,  # attribute id
              'value_ids': [(4, 2), ]},  # [(unknown ???, value_ids)]
             ),

            ]
        """
        # ********************

        product_id = self._models_objects.execute_kw(
            self._db, self._uid, self._password,
            'product.template', 'create',
            [{
                'name': sname+'29',
                'price': item['price'],
                'categ_id': 6,  # All / Можна продавати / Physical
                # 'default_code': '1111',
                'public_categ_ids': [[6, 0, [cat_id]]],
                # 'description_sale': 'super_puper_long',
                'website_description': item['desc'],
                'website_published': True,
                'image': item['image'],
                'attribute_line_ids': attrs_lines,
            }]
        )

        out('[id: %i] [%s] ' % (product_id, sname,))

    def set_attributes_for_item(self, id_item, attributes_list):
        id_item = 42
        attributes_list = [
            [{'attribute_id': 7,
              'values_ids': [9]}],
            # [],
        ]

        product_id = self._models_objects.execute_kw(
            self._db, self._uid, self._password,
            'product.template',
            'write',
            [65,
             {'attribute_line_ids': [
                [
                    0,
                    False,
                    {'attribute_id': 1,
                     'value_ids': [
                                6,
                                False,
                                [1]
                            ]
                     }
                 ],
             ]}],
            # [],
            #    'attribute_line_ids': attributes_list,
        )

        pass

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


# region 'conditions list const'
CON_IGNORE_SNAME_STARTS = ['Тканини для асортименту "ТИСА-МЕБЛІ"',
                           ]

CON_IGNORE_ATTRS = ['Гарантийный срок',
                    'Ручки для переноса',
                    'Тип подъемного механизма',
                    'Количество зон жесткости матраса',
                    'Количество спальных мест',
                    'Вид кровати',
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

        item = {}

        error_attrbitutes_values = {}
        attrs_list_names = {}

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

            # index
            item['index'] = index

            # name
            sname = str(row[1].value).strip() \
                .replace('"', '') \
                .replace('  ', ' ') \
                .replace('*', 'x')
            item['sname'] = sname
            out('[%i/%i] %s' % (index, total_count, sname))

            v = row[5].value
            item['price'] = int(v) if not (v is None) else 0

            v = row[3].value
            item['description'] = str(v) if not (v is None) else ''

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
                .replace('Столи гостьові-трансформери', 'Столи журнальні') \
                .replace('Стільці', 'Стільці та табурети') \
                .replace('Дитячі дивани', 'Дивани') \
                .replace('Кутові дивани', 'Дивани') \
                .replace('Прямі дивани', 'Дивани') \
                .strip()

            item['cat'] = cat
            cats[cat] = None

            # brand
            tmp = str(row[24].value).strip()
            tmp = tmp \
                .replace('Скиф', 'Скіф') \
                .replace('Тиса мебель', 'Тиса меблі') \
                .replace('Елисеевская мебель', 'Єлисеївські меблі') \
                .replace('Микс мебель', 'Мікс меблі') \
                .replace('Мелитополь мебель', 'Мелітополь меблі') \
                .replace('Еврокнижка', 'Єврокнижка') \
                .replace('деревянные ламели', 'букові ламелі')

            if tmp == '':
                tmp = 'Константа'
            extra_attr_dict['Бренд'] = [sname, tmp, 'str', '']

            # country
            tmp = str(row[26].value).strip()
            tmp = tmp.replace('Украина', 'Україна')
            if tmp == '':
                tmp = 'Україна'
            extra_attr_dict['Країна виробник'] = [sname, tmp, 'str', '']

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
                extra_attr_dict['Гарантія'] = [sname, '18', 'int', 'міc']

            # ignore some "with start" sname products
            for e in CON_IGNORE_SNAME_STARTS:
                if sname.startswith(e):
                    sname = ''
            if sname == '':
                continue

            # if sname.startswith('Тканини для асортименту "ТИСА-МЕБЛІ"'):
            #    continue

            # af = []
            # get list of tuples(name_product, name_attr, name_counter, value)
            # use step by 1

            # region 'convert work'
            attrs_list0 = [[sname,
                            str(val.value).strip(),
                            str(some_list[idx + 1].value).strip(),
                            str(some_list[idx + 2].value).strip()
                            .replace('.0', '')
                            .replace('*', 'x')
                            ]
                           for idx, val in enumerate(some_list)
                           if idx % 3 == 0
                           and val.value is not None]

            # skip ignored attributes
            attrs_list1 = [e for e in attrs_list0
                           if e[1] not in CON_IGNORE_ATTRS
                           and not (e[1] == 'Цвет' and e[3] == 'Разные цвета')
                           and not (e[1] == 'Розмір' and e[3] == '7000')
                           and not (e[1] == 'Состояние' and e[3] == 'Новое')
                           and not (e[1] == 'Тип' and e[3] == 'Для сна')
                           and not (e[1] == 'Цвет' and e[3] == 'Белый')
                           and not (e[1] == 'Цвет' and e[3] == 'Разные цвета')
                           and not (e[1] == 'Цвет обивки' and e[3] == 'Разные цвета')
                           and not (e[1] == 'Тип крепления к матрасу' and e[3] == 'четыре резинки по углам')
                           ]

            for e in attrs_list1:
                attrs_list_names[e[1]] = ''

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
                                     'Высота стола',
                                     'Длина стола в раздвинутом (разложенном) состоянии',
                                     'Длина стола в сдвинутом (сложенном) состоянии',
                                     'Ширина',
                                     'Ширина стола',
                                     'Глубина',
                                     ]:
                    e[3] = str(int(e[3]) / 10).replace('.0', '')
                    e[2] = 'см'

                if e[2] in ['см', 'кг', 'шт.']:
                    e[2] = 'int'

                if e[3] in ['да', 'нет', 'True', 'False']:
                    e[2] = 'bool'

                    e[3] = e[3].replace('да', 'так').replace('нет', 'ні').replace('True', 'так').replace('False', 'ні')

                # else string
                # if e[2] == '':
                #    e[2] = 'str'

                # attrs_list2.append(result)

            attrs_list = attrs_list1

            # add extra data
            for e in extra_attr_dict:
                attrs_list.append([sname, e, extra_attr_dict[e][0], extra_attr_dict[e][1]])
                # attrs_list.append([e, extra_attr_dict[e][1]])

            # item['attributes'] = attrs_list
            item['attributes'] = {}
            for e in attrs_list:
                # skip empty attribute values
                if not e[3].strip() == '':  # !!!!!!!!!!!!!!!
                    item['attributes'][e[1]] = e[3]
                    # test anomality attributes values !!!
                    if e[3] in [
                        '', 'г', '7', 'Мікс мелі', '309',
                        '890', '760', '4.7', '4.5', '7.6', 'False', 'True',
                        'Двоярусна', 'Ножки', 'Tik-Tak', 'взаимозаменяемый', 'двуспальная', 'левый',
                        'одноярусная кровать',
                        'шок',
                        # '101', '102', '103', '104', '105', '106', '107', '108',
                        ]:
                        error_attrbitutes_values[e[0] + ', ' + e[1]+'='+e[3]] = ''
                        out('      ******     Error attribute value: %s      *********' % e[3])

                    # out('empty attribute value')
            # detect types of data

            # endregion

            # OUTPUT
            for e in attrs_list:
                s = '%s\t%s\t%s\t%s\n' % (e[0], e[1], e[2], e[3])
                self._csv1.writelines(s)

                if e[3].strip() == '':
                    a = 1
                # just for debug
                attr_dict[e[3]] = ''
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

        # attrs_list_names
        out('\n\n==Attr_names_dict==')
        for e in sorted(attrs_list_names):
            out(e)

        out('\n\n==Attr_values_dict==')
        for e in sorted(attr_dict):
            out(e)

        out('\n\n==Error: error_attrbitutes_values==')
        for e in sorted(error_attrbitutes_values):
            out(e)


# c = Impxls()
# c.handle('export-products.xlsx')

# test
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

    {'name': 'техас',
     'price': 66,
     'cat_name': 'диван',
     'desc': 'super duper divan2',
     # 'cat_name': 'диван',
     'image': '',
     },

]

im = ImportToOdd(env('host'), env('db'), env('user'), env('pwd'))
# im.connect()

# im.test()
# im.unlink_attributes()
# im.create_attribute('форма', 'квадрат')

# im.set_attributes_for_item(None, None)
# cats = im.create_categories(list_1)

# im.unlink_item()

c = Impxls()
c.handle('export-products.xlsx')

exit(0)


out('\n==products==')
i = 0
for e in list_2:
    i += 1
    out('[%i/%i] [%s]' % (i, len(list_2), e['name'],))
    im.create_item(e)
