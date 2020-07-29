import os
import sys

import PDF2Word
from PyPDF2 import PdfFileReader, PdfFileWriter
from PyPDF2.pdf import PageObject
from win32com import client

from util_main import log, cross_iter, is_purchase_page

def docx2pdf(doc_fn, pdf_fn):
    log.info("Converting: %s --> %s" %(os.path.basename(doc_fn), os.path.basename(pdf_fn)))
    if os.path.exists(pdf_fn):
        log.info("File exist and skip.")
        return
    word = client.DispatchEx("Word.Application")
    doc_trans = word.Documents.Open(doc_fn)
    doc_trans.SaveAs(pdf_fn, FileFormat=17)
    doc_trans.Close()


def remove_first_page(pdf_fn, dst_fn):
    log.info("Converting: %s --> %s" %(os.path.basename(pdf_fn), os.path.basename(dst_fn)))
    if os.path.exists(dst_fn):
        log.info("File Exist, skip...")
        return dst_fn

    infile = PdfFileReader(pdf_fn, 'rb')
    output = PdfFileWriter()

    skip_pages = [0]
    for i in range(infile.getNumPages()):
        if i in skip_pages:
            continue
        p = infile.getPage(i)
        output.addPage(p)

    with open(dst_fn, 'wb') as f:
        output.write(f)
    log.info("Done.")
    return dst_fn

def convert_pdf_to_docx_v1(pdf_fn, word_fn):
    log.info("Converting: %s --> %s" % (os.path.basename(pdf_fn), os.path.basename(word_fn)))
    if os.path.exists(word_fn):
        log.info("File Exist, skip...")
        return
    pdf2word = PDF2Word.PDF2Word()
    try:
        pdf2word.setOutputDocumentFormat(PDF2Word.optOutputDocumentFormat.OPT_OUTPUT_DOCX)
        pdf2word.setConnectHyphens(True)
        pdf2word.setShrinkCharacterSpacingToPreventWrap(True)
        pdf2word.setFileConversionTimeout(6000000)
        pdf2word.ConvertToWord(pdf_fn, word_fn, "", 0, -1)
    except PDF2Word.PDF2WordException as ex:
        log.info(ex)
        sys.exit()

# 提供的3种模式：Absolutely / Reflow / without structure
# 均会破坏文档结构，导致生成的PDF结构混乱和文字重叠
def convert_pdf_to_docx_v2(inputFileName, outputFileName):
    log.info("Converting: %s --> %s" % (os.path.basename(inputFileName), os.path.basename(outputFileName)))
    if os.path.exists(outputFileName):
        log.info("File Exist, skip...")
        return
    pdf2word = PDF2Word.PDF2Word()
    try:
        pdf2word.setConversionMethod(PDF2Word.optConversionMethod.CNV_METHOD_USE_TEXTBOXES)
        pdf2word.setOutputDocumentFormat(PDF2Word.optOutputDocumentFormat.OPT_OUTPUT_DOCX_VIA_OFFICE)
        pdf2word.setDocumentType(PDF2Word.optDocumentType.DOCTYPE_MULTI_COLUMN)
        pdf2word.setAdjustSpacing(True)
        pdf2word.ConvertToWord(inputFileName, outputFileName, "", 0, -1)
    except PDF2Word.PDF2WordException as ex:
        log.info(ex)
        sys.exit()

def get_page_from_nums(pdf, page_nums):
    the_pages = []
    for n in page_nums:
        if n >= 0:
            page = pdf.getPage(n)
        elif n == -1:
            page = PageObject.createBlankPage(pdf)
        the_pages.append(page)
    return the_pages

def merge_pages(pdfTransFn1, pdfTransFn2, dst_fn, page_nums1=None, page_nums2=None):
    infile1 = PdfFileReader(pdfTransFn1, 'rb')
    infile2 = PdfFileReader(pdfTransFn2, 'rb')
    if page_nums1 is None:
        page_nums1 = [("first", [i]) for i in range(infile1.getNumPages())]
    if page_nums2 is None:
        page_nums2 = [("second", [i]) for i in range(infile2.getNumPages())]
    pages_sides = cross_iter(page_nums1, page_nums2)

    log.info("Merge Pages: %d+%d --> %s" % (len(page_nums1), len(page_nums2), dst_fn))
    # if os.path.exists(dst_fn):
    #     log.info("File Exist and skip.")
    #     return

    output = PdfFileWriter()
    for (side, page_nums) in pages_sides:
        if side == "first":
            pages = get_page_from_nums(infile1, page_nums)
        elif side == "second":
            pages = get_page_from_nums(infile2, page_nums)
        else:
            raise Exception("Fuck!")
        for p in pages:
            output.addPage(p)

    # 只考虑头个文件的书签
    # bookmarks = infile1.getOutlines()
    # if len(bookmarks) > 0:
    #     for bm in bookmarks:
    #         bm_page_num = infile1.getDestinationPageNumber(bm)

    # 利用python处理pdf：奇数页pdf末尾添加一个空白页 - https://zhuanlan.zhihu.com/p/34246341
    # 青梅煮马: 刚好遇到这个问题 把 PyPDF2\utils.py 第238行的'latin-1'编码修改为'uft-8'即可
    # 上面的方法会影响兼容性，改下面的方法
    # PyPDF2 编码问题’latin-1′ codec can’t encode characters in position 8-11: ordinal not in range(256)  https://www.codenong.com/cs105218309/
    with open(dst_fn, 'wb') as f:
        output.write(f)

def extend_to_odd(l, fill=-1):
    the_len = len(l)
    if the_len == 0:
        raise ValueError("Input was empty: %s" %(repr(l)))
    if the_len % 2 == 0:
        l.append(fill)

def get_merged_pdf_nums(side_nums1, side_nums2):
    first_added = 0
    second_added = 0
    merged_nums = []
    for side, page_nums in cross_iter(side_nums1, side_nums2):
        if side == "first":
            cur_nums = [num + second_added if num >= 0 else num for num in page_nums]
            first_added += len(page_nums)
        elif side == "second":
            cur_nums = [num + first_added if num >= 0 else num for num in page_nums]
            second_added += len(page_nums)
        merged_nums.append(cur_nums)
    final_nums = [('second', nums) for nums in merged_nums]
    return final_nums

def get_content_pages(pdf_fn, side, mode="1To1"):
    infile = PdfFileReader(pdf_fn, "rb")
    page_count = infile.getNumPages()
    is_content_page = True
    purchase_count = 0
    page_numbers = []
    pages = []

    if mode == "1To1":
        for i in range(page_count):
            cur_page = infile.getPage(i)
            if is_content_page:
                page_numbers.append((side, [i]))
                is_content_page = False
                pages.append(cur_page)
            else:
                if is_purchase_page(cur_page):
                    is_content_page = True
                    purchase_count += 1
                else:
                    pass
        print("Content Page Count:\t%d" % (len(page_numbers)))
    elif mode == "1ToN":
        the_batch_pages = []
        for i in range(page_count):
            cur_page = infile.getPage(i)
            if is_purchase_page(cur_page):
                extend_to_odd(the_batch_pages)
                page_numbers.append((side, the_batch_pages))
                the_batch_pages = []
                purchase_count += 1
            else:
                the_batch_pages.append(i)
        if len(the_batch_pages) > 0:
            extend_to_odd(the_batch_pages)
            page_numbers.append((side, the_batch_pages))
        print("Content Page Count:\t%d" % (len(page_numbers)))
        print("Flatten Page Count (1toN):\t%d" %(sum([len(l) for l in page_numbers])))

    print("Purchase Count:\t\t%d" % (purchase_count))
    return page_numbers

convert_pdf_to_docx = convert_pdf_to_docx_v1

