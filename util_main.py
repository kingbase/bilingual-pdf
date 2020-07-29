import logging
import os
import sys
from collections import Iterable

from PyPDF2.pdf import PageObject
from docx.text.paragraph import Paragraph

PURCHASE_PAGE_TEXT1 = "www.pdfonline.com/purchase"
PURCHASE_PAGE_TEXT2 = "info@bcltechnologies"
def is_purchase_page(page):
    assert isinstance(page, PageObject)
    text = page.extractText()
    if PURCHASE_PAGE_TEXT1 in text or \
            PURCHASE_PAGE_TEXT2 in text:
        return True
    return False

# How to extract text from an existing docx file using python-docx
# https://stackoverflow.com/questions/25228106/
def para2text(p):
    assert isinstance(p, Paragraph)
    rs = p._element.xpath('.//w:t')
    return u" ".join([r.text for r in rs if r.text is not None])

# a=[1,3,5], b=[2,4], result=[1,2,3,4,5]
def cross_iter(list_a, list_b):
    iter_a = iter(list_a)
    iter_b = iter(list_b)
    while 1:
        ele_a = next(iter_a, None)
        ele_b = next(iter_b, None)
        if ele_a is not None:
            yield ele_a
        if ele_b is not None:
            yield ele_b
        if ele_a is None and ele_b is None:
            break

def batch(iterable, size=1):
    '''
    将迭代器批次化，以应对无法一次性完成的情况
    '''
    if not isinstance(iterable, Iterable):
        raise ValueError("Input should be iterable")
    l = len(iterable)
    for ndx in range(0, l, size):
        yield iterable[ndx:min(ndx + size, l)]

def check_exists(fn):
    if not os.path.exists(fn):
        sys.exit("File Not Exists: %s" %(fn))

def check_file_exists(fn):
    check_exists(fn)
    if not os.path.isfile(fn):
        sys.exit("Name is Not File: %s" %(fn))

def check_dir_exists(fn):
    check_exists(fn)
    if not os.path.isdir(fn):
        sys.exit("Name is Not Dir: %s" %(fn))

log         = logging.getLogger("BilingualPdf")
std_handler = logging.StreamHandler(sys.stdout)
FORMATTER   = logging.Formatter("[%(asctime)s]%(funcName)s@%(lineno)d: %(message)s", "%H:%M:%S")
std_handler.setFormatter(FORMATTER)
log.addHandler(std_handler)
log.setLevel(logging.DEBUG)
