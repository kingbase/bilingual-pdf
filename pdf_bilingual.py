#encoding: utf-8
import sys
import os.path
from shutil import copyfile

import diskcache as dc

from util_convert import remove_first_page, convert_pdf_to_docx, docx2pdf, merge_pages, \
    get_content_pages, get_merged_pdf_nums
from util_trans import translate_doc
from util_main import cross_iter, check_dir_exists, check_file_exists, log

if len(sys.argv) < 4:
   print("Please specify mode / input file name / tmp output folder / [translated docx](optional).")
   sys.exit()

SUPPORTED_MODES = ["Single1To1", "Single1ToN", "KnownDoc"]
mode          = sys.argv[1]
inputFileName = sys.argv[2]
outputDirName = sys.argv[3]
assert inputFileName.endswith(".pdf")
check_file_exists(inputFileName)
check_dir_exists(outputDirName)
cache_trans = dc.Cache(outputDirName)
if mode not in SUPPORTED_MODES:
    print("Supported Modes: %s" %(",".join(SUPPORTED_MODES)))
    print("Check help for how to choose modes.")
    sys.exit()

# 第0步：将原始PDF复制到目标目录
inputBaseNameWithoutExt = os.path.basename(inputFileName).replace(".pdf", "")

if mode.lower().startswith("single"):
    # 第1步：将PDF分为2个，以应对该共享软件每次间隔1页留空白，方法是将PDF的第1页删掉
    pdfFileName1 = os.path.join(outputDirName, "%s_1_RawPart1.pdf" %(inputBaseNameWithoutExt))
    pdfFileName2 = os.path.join(outputDirName, "%s_1_RawPart2.pdf" % (inputBaseNameWithoutExt))
    log.info("%s, %s" %(pdfFileName1, pdfFileName2))
    if not os.path.exists(pdfFileName1):
        copyfile(inputFileName, pdfFileName1)
    remove_first_page(pdfFileName1, pdfFileName2)

    # 第2步：分别将这2个PDF转为DOCX
    docFileName1 = os.path.join(outputDirName, "%s_2_RawPart1.docx" %(inputBaseNameWithoutExt))
    docFileName2 = os.path.join(outputDirName, "%s_2_RawPart2.docx" %(inputBaseNameWithoutExt))
    log.info("%s, %s" %(docFileName1, docFileName2))
    convert_pdf_to_docx(pdfFileName1, docFileName1)
    convert_pdf_to_docx(pdfFileName2, docFileName2)

    # 第3步：将这2个DOCX翻译，仍然保存为DOCX
    docTransFn1 = os.path.join(outputDirName, "%s_3_TranslatedPart1.docx" %(inputBaseNameWithoutExt))
    docTransFn2 = os.path.join(outputDirName, "%s_3_TranslatedPart2.docx" %(inputBaseNameWithoutExt))
    log.info("%s, %s" %(docTransFn1, docTransFn2))
    translate_doc(docFileName1, docTransFn1, cache_trans)
    translate_doc(docFileName2, docTransFn2, cache_trans)

    # 第4步：将这2个DOCX转换为PDF
    pdfTransFn1 = os.path.join(outputDirName, "%s_4_TranslatedPart1.pdf" %(inputBaseNameWithoutExt))
    pdfTransFn2 = os.path.join(outputDirName, "%s_4_TranslatedPart2.pdf" %(inputBaseNameWithoutExt))
    log.info("%s, %s" %(pdfTransFn1, pdfTransFn2))
    docx2pdf(docTransFn1, pdfTransFn1)
    docx2pdf(docTransFn2, pdfTransFn2)

    # 第5步：将这2个PDF合并为1个PDF，即为最终输出
    # 这一步可能会导致信息缺失，因为个别页面可能在原始的PDF中只有1页
    # 但在转换后有1页多，此时我们只取第1页，以达到双页并排显示的效果
    if mode.endswith("1To1"):
        pdfPageNumbers1 = get_content_pages(pdfTransFn1, "first")
        pdfPageNumbers2 = get_content_pages(pdfTransFn2, "second")
    elif mode.endswith("1ToN"):
        pdfPageNumbers1 = get_content_pages(pdfTransFn1, "first", "1ToN")
        pdfPageNumbers2 = get_content_pages(pdfTransFn2, "second", "1ToN")
    mergedTransPdfFn = os.path.join(outputDirName, "%s_5_TranslatedAll.pdf" %(inputBaseNameWithoutExt))
    merge_pages(pdfTransFn1, pdfTransFn2, mergedTransPdfFn, pdfPageNumbers1, pdfPageNumbers2)
    finalFileName = os.path.join(outputDirName, "%s_6_FINAL.pdf" %(inputBaseNameWithoutExt))
    if mode.endswith("1To1"):
        merge_pages(inputFileName, mergedTransPdfFn, finalFileName)
    elif mode.endswith("1ToN"):
        pdfPageNumbersFinal = get_merged_pdf_nums(pdfPageNumbers1, pdfPageNumbers2)
        merge_pages(inputFileName, mergedTransPdfFn, finalFileName, None, pdfPageNumbersFinal)
elif mode == "KnownDoc":
    toCopyFileName = os.path.join(outputDirName, "%s_1_Raw.pdf" %(inputBaseNameWithoutExt))
    pdfFileName = toCopyFileName
    convertedDocFileName = sys.argv[4]
    translatedDocFileName = os.path.join(outputDirName, "%s_2_Translated.docx" %(inputBaseNameWithoutExt))
    translate_doc(convertedDocFileName, translatedDocFileName, cache_trans)
    translatedPdfFileName = os.path.join(outputDirName, "%s_3_Translated.pdf" %(inputBaseNameWithoutExt))
    docx2pdf(translatedDocFileName, translatedPdfFileName)
    finalFileName = os.path.join(outputDirName, "%s_4_FINAL.pdf" %(inputBaseNameWithoutExt))
    merge_pages(inputFileName, translatedPdfFileName, finalFileName)
log.info("Convert Success: %s" %(finalFileName))