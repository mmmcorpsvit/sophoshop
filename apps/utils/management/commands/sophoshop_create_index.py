import logging

from django.core.management.base import BaseCommand, CommandError
from oscar.core.loading import get_class, get_classes, get_model
from elasticsearch import Elasticsearch

ImportingError = get_class('partner.exceptions', 'ImportingError')
Partner, StockRecord = get_classes('partner.models', ['Partner', 'StockRecord'])
ProductClass, ProductAttribute, Product, Category, ProductCategory = get_classes(
    'catalogue.models', ('ProductClass', 'ProductAttribute', 'Product',
                         'Category', 'ProductCategory'))

AttributeOption, AttributeOptionGroup = get_classes(
    'catalogue.models', ('AttributeOption', 'AttributeOptionGroup'))

ProductImage = get_model('catalogue', 'productimage')


logger = logging.getLogger('oscar.catalogue.import')


# https://www.youtube.com/watch?v=GELah8on52k
# 23:25

class Command(BaseCommand):

    def handle(self, *args, **options):
        es = Elasticsearch()
        # es.create('sophoshop')

        count = 0
        # for p in Product.objects.all().order_by('id')[4000:]:
        for p in Product.objects.all().order_by('id'):
            a = es.search(index='products', body={"query": {"match": {"product_id": p.id}}})
            
            if p.category:
                slug = p.category.slug
            else:
                continue
            if p.image:
                image = p.image.url
            else:
                image = ''
            count += 1
            doc = {
                'name': p.name,
                'product_id': p.id,
                'price': p.price,
                'title': p.title,
                'image': image,
                'slug': p.slug,
                'created_at': p.created_at,
                'updated_at': p.updated_at,
            }
            filters = p.get_filters()
            for key in filters:
                doc[key] = filters[key]
            
            res = es.index(index='products', doc_type=slug,  body=doc)
            logging.info(res)

            logging.info("***********************")
            logging.info("prod id %s" % p.id)
            logging.info("count %s" % count)
            logging.info("***********************")
