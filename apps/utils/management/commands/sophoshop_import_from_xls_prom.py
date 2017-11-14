# import os
# import tarfile
# import zipfile
# import zlib

import shutil
import tempfile
import urllib3

import logging

from django.core.files import File
from django.core.management.base import BaseCommand, CommandError
# from oscar.core.loading import get_class

# from decimal import Decimal as D

# from django.db.transaction import atomic
# from django.db import transaction
# from django.utils.translation import ugettext_lazy as _

from oscar.apps.catalogue.categories import create_from_breadcrumbs
from oscar.core.loading import get_class, get_classes, get_model

from openpyxl import load_workbook

ImportingError = get_class('partner.exceptions', 'ImportingError')
Partner, StockRecord = get_classes('partner.models', ['Partner', 'StockRecord'])
ProductClass, Product, Category, ProductCategory = get_classes(
    'catalogue.models', ('ProductClass', 'Product', 'Category', 'ProductCategory'))
ProductImage = get_model('catalogue', 'productimage')


logger = logging.getLogger('oscar.catalogue.import')
urllib3.disable_warnings()

# use: sophoshop_import_from_xls_prom export-products.xlsx --flush --add_images


def download(c, url, filename):
    """
    Download file
    :param c: c = urllib3.PoolManager()
    :param url: URL
    :param filename: filename where file saved
    # :return: data of file
    """

    # logger.error(url)
    # logger.error(filename)

    with c.request('GET', url, preload_content=False) as resp, open(filename, 'wb') as out_file:
        shutil.copyfileobj(resp, out_file)

    resp.release_conn()  # not 100% sure this is required though
    # return data


class Impxls(object):
    _flush = False
    _add_images = False

    def __init__(self, logger2, flush=False, add_images=False):
        self.logger = logger2
        self._flush = flush
        self._add_images = add_images

    def _create_item(self, product_class, category_str, upc, title, description, images_urls, price, stats):
        # Ignore any entries that are NULL
        if description == 'NULL':
            description = ''

        # Create item class and item
        product_class, __ = ProductClass.objects.get_or_create(name=product_class)
        try:
            item = Product.objects.get(upc=upc)
            stats['updated_items'] += 1
        except Product.DoesNotExist:
            item = Product()
            stats['new_items'] += 1
        item.upc = upc
        item.title = title
        item.description = description
        item.product_class = product_class
        if not (price is None):
            item.price = price
        item.save()

        # Category
        cat = create_from_breadcrumbs(category_str)
        ProductCategory.objects.update_or_create(product=item, category=cat)

        c = urllib3.PoolManager()

        # image
        if self._add_images:
            images = str(images_urls).split(',')
            for image in images:
                image_url = image.strip()
                if len(image_url) < 5:
                    continue
                # logger.info('download image: %s' % image)

                # data = download(image)

                file_name = image.replace('https://images.ua.prom.st/', '')
                fn = tempfile.gettempdir() + '\\' + file_name

                download(c, image_url, fn)

                new_file = File(open(fn, 'rb'))
                im = ProductImage(product=item)
                im.original.save(file_name, new_file, save=False)
                im.save()
                logger.debug('Image added to "%s"' % item)

        # stockrecord
        self._create_stockrecord(item, 'Світ Комфорту', upc,
                                 price, 14)

        return item

    @staticmethod
    def _create_stockrecord(item, partner_name, partner_sku,
                            price, num_in_stock):
        def d(x):
            return int(x)

        # Create partner and stock record
        partner, _ = Partner.objects.get_or_create(name=partner_name)
        try:
            stock = StockRecord.objects.get(partner_sku=partner_sku)
        except StockRecord.DoesNotExist:
            stock = StockRecord()

        stock.product = item
        stock.partner = partner
        stock.partner_sku = partner_sku
        stock.price_excl_tax = d(price)
        stock.price_retail = d(price)
        stock.num_in_stock = num_in_stock
        stock.save()

    def _flush_product_data(self):
        """Flush out product and stock models"""
        logger.info('Flush start')
        ProductCategory.objects.all().delete()
        Category.objects.all().delete()
        Product.objects.all().delete()
        ProductClass.objects.all().delete()
        Partner.objects.all().delete()
        StockRecord.objects.all().delete()
        if not self._add_images:
            logger.info('Flush images')
            ProductImage.objects.all().delete()
        logger.info('Flush end')

    def handle(self, fn):
        if self._flush:
            self._flush_product_data()

        stats = {'new_items': 0,
                 'updated_items': 0}

        wb2 = load_workbook(fn)
        logger.info('Spread sheats names: %s' % wb2.get_sheet_names())
        wb = wb2.worksheets[0]

        skip_first_row = True
        index = 0
        cats = dict()

        for row in wb.rows:
            index += 1
            if skip_first_row:
                logger.info('[0 = skip]')
                skip_first_row = False
                continue

            # logger.info(skip_first_row)
            v = row[5].value
            price = int(v) if not (v is None) else 0

            v = row[3].value
            description = str(v) if not (v is None) else ''

            cat = str(row[15].value) \
                .replace('Матраци Sleep&Fly', 'Матраци') \
                .replace('Матраци Evolution', 'Матраци') \
                .replace('Матраци Sleep&fly Organic', 'Матраци') \
                .replace('Матраци Take&go Bamboо', 'Матраци') \
                .replace('Матраци Sleep&fly uno', 'Матраци') \
                .replace('Матраци на дивани', 'Футони і топери') \
                .replace('Наматрацникии', 'Наматрацники і підматрацники') \
                .replace("Дерев'яні ліжка", 'Ліжка') \
                .replace("Дитячі ліжка", 'Ліжка') \
                .replace('Столи', 'Столи гостьові') \
                .replace('Столи гостьові-трансформери', 'Столи журнальні')\
                .replace('Стільці', 'Стільці та табурети')

            cats[cat] = None

            self._create_item(
                product_class=str(row[16].value).replace('https://prom.ua/', ''),
                category_str=cat,
                upc=str(row[20].value),
                title=str(row[1].value).strip(),
                description=description,
                price=price,
                stats=stats,
                images_urls=str(row[11].value),
            )

            # logger.info('[%i/%i] %s ' % (row.index, wb.rows.count, row[1].value))
            logger.info('[%i/%i] [%s] %s' % (index - 1, wb.max_row, cat, row[1].value,))
            # logger.info(row[1].value)

        msg = "New items: %d, updated items: %d" % (stats['new_items'],
                                                    stats['updated_items'])
        self.logger.info(msg)
        self.logger.info(cats)


class Command(BaseCommand):
    help = 'Import Products and Categories from Prom.ua exported XLS file'

    def add_arguments(self, parser):
        parser.add_argument(
            'filename', nargs='+',
            help='/path/to/file1.xls /path/to/file2.xls ...')
        parser.add_argument(
            '--flush',
            action='store_true',
            dest='flush',
            default=False,
            help='Flush tables before importing')

        parser.add_argument(
            '--add_images',
            action='store_true',
            dest='add_images',
            default=False,
            help='Process images importing')

    def handle(self, *args, **options):
        logger.info("Starting catalogue import")

        for file_path in options['filename']:
            logger.info(" - Importing records from '%s'" % file_path)
            try:
                xls = Impxls(logger, flush=options.get('flush'), add_images=options.get('add_images'))
                xls.handle(file_path)

            except ImportingError as e:
                raise CommandError(str(e))
