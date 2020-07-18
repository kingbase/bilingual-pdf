## 功能介绍

也许你的英文能力尚可，但作为中国人，自然是母语的阅读速度更快。

有时会遇到长篇的英文文档，读起来很费劲，如果能够用中文辅助以加快阅读速度，自然是最妙不过。

本 repo 就是用于这样的目的：

对原始的 PDF ，会在其下一页翻译出对应的中文页面，以加速您的阅读速度：不重要的内容，用中文加速阅读。重要的内容，看原文。

效果如图：

![Think Complexity 2 Bilingual Demo](https://github.com/kingbase/bilingual-pdf/raw/master/images/thinkcomplexity2_demo.png)

## 安装说明

本 repo 只能跑在 **Windows** 上，这是由于以下2个限制：

PDF --> Word，使用的是 pdf2word 提供的 SDK，其提供了 Windows 和 Linux 版，我只在 Windows 上运行了。

Word --> PDF，依赖于 Office Word 的 COM 接口。

### 安装步骤

1. 安装 pdf2word converter sdk
  1. 需要在 https://www.pdfonline.com/downloads/easyconvertersdk/demos_easyconvertersdk-word-2.asp 申请该SDK的试用。
  2. 安装后，将安装目录添加到 PYTHONPATH ，假定安装位置是 `C:\Program Files\BCL Technologies\easyConverter SDK 5\Rtf`，则在 Windows 上需要设置为：`setx PYTHONPATH "%PYTHONPATH%;C:\Program Files\BCL Technologies\easyConverter SDK 5\Rtf"`
2. 安装其他依赖 `pip install -r requirements.txt`

## 使用说明

限制：只能处理可编辑的PDF，无法处理扫描版。

由于不同 PDF 的转换难度不同，需要一定程度的人工介入。

不同模式的区别

- Single1To1，适用于页面排版比较简单的PDF，如常见的英文书。
- Single1ToN，适用于页面排版比较复杂的PDF，如常见的学术论文中的双栏排版。
- KnownDoc，你希望自己完成 PDF --> Word 的步骤，仅保留 Word --> PDF --> Merge PDF 。

## 用法示例

我们以 http://greenteapress.com/complexity2/thinkcomplexity2.pdf 为例。

假设下载到本地为：`D:\down\thinkcomplexity2.pdf`，使用的中间目录为`D:\work\tmp`

`python pdf_bilingual.py Single1To1 "D:\down\thinkcomplexity2.pdf" d:\work\tmp`
