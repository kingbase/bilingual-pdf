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
mode       = sys.argv[1]
input_path = sys.argv[2]
output_dir = sys.argv[3]
assert input_path.endswith(".pdf")
check_file_exists(input_path)
check_dir_exists(output_dir)
cache_trans = dc.Cache(output_dir)
if mode not in SUPPORTED_MODES:
    print("Supported Modes: %s" %(",".join(SUPPORTED_MODES)))
    print("Check help for how to choose modes.")
    sys.exit()

# 第0步：将原始PDF复制到目标目录
input_fn_without_ext = os.path.basename(input_path).replace(".pdf", "")

if mode.lower().startswith("single"):
    # 第1步：将PDF分为2个，以应对该共享软件每次间隔1页留空白，方法是将PDF的第1页删掉
    pdf_fn1 = os.path.join(output_dir, "%s_1_RawPart1.pdf" % (input_fn_without_ext))
    pdf_fn2 = os.path.join(output_dir, "%s_1_RawPart2.pdf" % (input_fn_without_ext))
    log.info("%s, %s" % (pdf_fn1, pdf_fn2))
    if not os.path.exists(pdf_fn1):
        copyfile(input_path, pdf_fn1)
    remove_first_page(pdf_fn1, pdf_fn2)

    # 第2步：分别将这2个PDF转为DOCX
    doc_fn1 = os.path.join(output_dir, "%s_2_RawPart1.docx" % (input_fn_without_ext))
    doc_fn2 = os.path.join(output_dir, "%s_2_RawPart2.docx" % (input_fn_without_ext))
    log.info("%s, %s" % (doc_fn1, doc_fn2))
    convert_pdf_to_docx(pdf_fn1, doc_fn1)
    convert_pdf_to_docx(pdf_fn2, doc_fn2)

    # 第3步：将这2个DOCX翻译，仍然保存为DOCX
    doc_trans_fn1 = os.path.join(output_dir, "%s_3_TranslatedPart1.docx" % (input_fn_without_ext))
    doc_trans_fn2 = os.path.join(output_dir, "%s_3_TranslatedPart2.docx" % (input_fn_without_ext))
    log.info("%s, %s" % (doc_trans_fn1, doc_trans_fn2))
    translate_doc(doc_fn1, doc_trans_fn1, cache_trans)
    translate_doc(doc_fn2, doc_trans_fn2, cache_trans)

    # 第4步：将这2个DOCX转换为PDF
    pdf_trans_fn1 = os.path.join(output_dir, "%s_4_TranslatedPart1.pdf" % (input_fn_without_ext))
    pdf_trans_fn2 = os.path.join(output_dir, "%s_4_TranslatedPart2.pdf" % (input_fn_without_ext))
    log.info("%s, %s" % (pdf_trans_fn1, pdf_trans_fn2))
    docx2pdf(doc_trans_fn1, pdf_trans_fn1)
    docx2pdf(doc_trans_fn2, pdf_trans_fn2)

    # 第5步：将这2个PDF合并为1个PDF，即为最终输出
    # 这一步可能会导致信息缺失，因为个别页面可能在原始的PDF中只有1页
    # 但在转换后有1页多，此时我们只取第1页，以达到双页并排显示的效果
    if mode.endswith("1To1"):
        pdf_page_numbers1 = get_content_pages(pdf_trans_fn1, "first")
        pdf_page_numbers2 = get_content_pages(pdf_trans_fn2, "second")
    elif mode.endswith("1ToN"):
        pdf_page_numbers1 = get_content_pages(pdf_trans_fn1, "first", "1ToN")
        pdf_page_numbers2 = get_content_pages(pdf_trans_fn2, "second", "1ToN")
    merged_trans_pdf_fn = os.path.join(output_dir, "%s_5_TranslatedAll.pdf" % (input_fn_without_ext))
    merge_pages(pdf_trans_fn1, pdf_trans_fn2, merged_trans_pdf_fn, pdf_page_numbers1, pdf_page_numbers2)
    final_fn = os.path.join(output_dir, "%s_6_FINAL.pdf" % (input_fn_without_ext))
    if mode.endswith("1To1"):
        merge_pages(input_path, merged_trans_pdf_fn, final_fn)
    elif mode.endswith("1ToN"):
        pdf_page_numbers_final = get_merged_pdf_nums(pdf_page_numbers1, pdf_page_numbers2)
        merge_pages(input_path, merged_trans_pdf_fn, final_fn, None, pdf_page_numbers_final)
elif mode == "KnownDoc":
    converted_doc_fn = sys.argv[4]
    translated_doc_fn = os.path.join(output_dir, "%s_2_Translated.docx" % (input_fn_without_ext))
    translate_doc(converted_doc_fn, translated_doc_fn, cache_trans)
    translated_pdf_fn = os.path.join(output_dir, "%s_3_Translated.pdf" % (input_fn_without_ext))
    docx2pdf(translated_doc_fn, translated_pdf_fn)
    final_fn = os.path.join(output_dir, "%s_4_FINAL.pdf" % (input_fn_without_ext))
    merge_pages(input_path, translated_pdf_fn, final_fn)
log.info("Convert Success: %s" % (final_fn))
