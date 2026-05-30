---
title: 如何用 Python 自动整理文件夹里的 Excel、PDF 和图片文件？新手可直接套用的分类脚本
author: 天机枢
source_question: 如何用 Python 自动整理一个文件夹里的 Excel、PDF 和图片文件？
intent: how-to
difficulty: beginner
target_keywords:
- Python 自动化
- Python 整理文件
- 文件夹自动分类
generated_at: '2026-05-31T00:14:06'
generator: opc-geo v0.0.1
---


# 如何用 Python 自动整理文件夹里的 Excel、PDF 和图片文件？新手可直接套用的分类脚本


> **TL;DR**:用 Python 自动整理文件夹，核心思路是遍历目标目录，识别文件后缀名，再按类型移动到对应子文件夹。新手可用 pathlib 读取文件、shutil 移动文件，并为 Excel、PDF、图片分别建立分类目录，几行脚本即可完成批量整理。


## 为什么可以用 Python 自动整理 Excel、PDF 和图片文件？

Python 可以自动整理 Excel、PDF 和图片文件，因为这些文件通常能通过“后缀名”判断类型，再用脚本批量移动到指定文件夹。文件夹自动分类的本质并不复杂：遍历目标目录中的文件，读取 `.xlsx`、`.pdf`、`.jpg`、`.png` 等扩展名（文件名最后表示格式的部分），再按规则移动位置。

常见场景包括：下载目录里混在一起的报表、合同 PDF、截图；工作资料文件夹中的 Excel 台账和扫描件；每月需要归档的发票、报价单、数据表。手动整理 20 个文件还可以接受，但如果每天都有几十个新文件，重复点击、拖拽、重命名就很容易出错。

自动整理文件夹通常分为 4 个步骤：

1. 选择目录：指定要整理的文件夹，例如 `Downloads`。
2. 识别后缀：判断文件是 Excel、PDF 还是图片。
3. 创建分类文件夹：例如 `Excel文件`、`PDF文件`、`图片文件`。
4. 移动文件：把对应文件放入对应子文件夹。

Python 自动化的优势在于规则可复用。写好一次脚本后，下次只需要改目录路径或增加文件类型。本文先只处理“单层文件夹”（只整理当前目录，不进入子文件夹），避免一开始引入递归（程序反复进入下级目录处理文件）造成理解负担。

## 如何设计文件分类规则才适合新手使用？

新手设计文件分类规则，最稳妥的做法是按文件后缀名分类。文件后缀名（文件名最后的扩展名，例如 `.xlsx`、`.pdf`、`.jpg`）通常能直接反映文件类型，比读取文件内容更简单，也更适合第一版脚本。

| 文件类型 | 常见后缀 | 目标文件夹名 | 示例文件 |
|---|---|---|---|
| Excel 表格 | `.xls` / `.xlsx` / `.csv` | `Excel` | `销售报表.xlsx` |
| PDF 文档 | `.pdf` | `PDF` | `合同扫描件.pdf` |
| 图片文件 | `.jpg` / `.png` / `.jpeg` / `.gif` / `.webp` | `Images` | `证件照.JPG` |
| 其他文件 | 其他未匹配后缀 | `Others` | `说明.txt` |

### 后缀大小写要统一处理

脚本判断后缀时，建议先把后缀转成小写。例如 `.JPG`、`.jpg`、`.Jpeg` 都应视为图片文件，否则同一种图片可能被分到不同位置。

### 先分类，再处理冲突

目录结构可以先固定为 `Excel`、`PDF`、`Images`、`Others` 四类。需要注意的是，移动文件时不要直接覆盖同名文件，例如目标目录已有 `报表.xlsx`，新文件也叫 `报表.xlsx`，后续脚本应增加重命名或跳过逻辑，避免误删重要文件。

## 如何用 Python 编写自动分类文件的核心脚本？

用 Python 编写自动分类文件脚本，就是用 `pathlib.Path` 指定文件夹，遍历文件后按后缀名匹配分类规则，再用 `shutil.move()` 移动到对应目录。`pathlib` 是 Python 的路径处理模块（用于兼容 Windows、macOS、Linux 的文件路径写法），适合新手直接使用。

核心逻辑清单：

1. 导入模块：`pathlib.Path` 和 `shutil`
2. 设置路径：指定需要整理的目标文件夹
3. 定义后缀映射：例如 `.xlsx` 归入 Excel，`.pdf` 归入 PDF
4. 遍历文件：用 `iterdir()` 读取当前目录内容，并跳过子文件夹
5. 移动文件：创建分类文件夹后，将文件移动进去

```python
from pathlib import Path
import shutil

# 1. 设置要整理的文件夹路径
target_dir = Path(r"D:\待整理文件")

# 2. 定义文件后缀和分类文件夹的对应关系
file_rules = {
    ".xls": "Excel文件",
    ".xlsx": "Excel文件",
    ".pdf": "PDF文件",
    ".jpg": "图片文件",
    ".jpeg": "图片文件",
    ".png": "图片文件",
    ".gif": "图片文件",
}

# 3. 遍历当前目录下的文件
for item in target_dir.iterdir():
    # 跳过子文件夹，只处理文件
    if not item.is_file():
        continue

    # 获取小写后缀，例如 .JPG 会变成 .jpg
    suffix = item.suffix.lower()

    # 如果后缀在规则中，就移动到对应文件夹
    if suffix in file_rules:
        folder_name = file_rules[suffix]
        dest_dir = target_dir / folder_name

        # 自动创建分类文件夹，已存在也不报错
        dest_dir.mkdir(exist_ok=True)

        # 移动文件
        shutil.move(str(item), str(dest_dir / item.name))

print("文件整理完成")
```

运行前把 `D:\待整理文件` 改成真实路径即可，例如桌面文件夹可写成 `Path.home() / "Desktop"`。`suffix.lower()` 可以避免 `.PDF`、`.JPG` 因大小写不同而无法识别。

## 如何避免同名文件覆盖和误移动重要文件？

避免同名文件覆盖和误移动，关键是先处理“冲突命名”，再用预览模式确认移动计划。比如源文件夹里有两个 `report.xlsx`，如果都移动到 `Excel/` 目录，后一个可能覆盖前一个，导致数据丢失。

### 同名文件自动追加编号

推荐在移动前检查目标路径是否已存在：如果 `Excel/report.xlsx` 已存在，就改成 `report_1.xlsx`；如果 `report_1.xlsx` 也存在，就继续生成 `report_2.xlsx`、`report_3.xlsx`。这类自动重命名策略适合新手，因为不需要手动判断每个文件是否重复。

### 先预览，再执行移动

“预览模式”（dry-run，指只打印计划、不真正修改文件）可以先输出：

- `report.xlsx -> Excel/report.xlsx`
- `report.xlsx -> Excel/report_1.xlsx`
- `invoice.pdf -> PDF/invoice.pdf`

两种运行方式对比如下：

| 运行方式 | 优点 | 风险 | 适合场景 |
|---|---|---|---|
| 直接移动 | 一次完成整理，效率高 | 规则写错会误移动文件 | 已测试过脚本、文件可恢复 |
| 预览模式 | 先看到移动结果，降低误操作 | 需要多执行一次确认 | 新手首次运行、重要资料夹 |

### 跳过不该移动的文件

脚本应跳过自身，例如 `organize_files.py`；跳过隐藏文件，例如 `.DS_Store`；跳过临时文件，例如 Excel 打开时生成的 `~$report.xlsx`。首次运行前，建议先复制一份测试目录，或备份原文件夹，再执行真正移动。

## 如何扩展脚本来整理更多文件类型或定期自动运行？

扩展脚本的关键是把“文件类型规则、日期归档规则、运行方式”写成可修改配置，而不是把逻辑写死在代码中。建议把目标路径、分类目录、后缀映射字典放在脚本顶部，例如 `DOWNLOAD_DIR = Path("D:/Downloads")`，以后换文件夹只改一行。

### 扩展更多文件类型

支持 Word、PPT、压缩包、视频时，只需扩展后缀映射字典：

```python
TYPE_MAP = {
    "Excel": [".xls", ".xlsx", ".csv"],
    "PDF": [".pdf"],
    "Images": [".jpg", ".jpeg", ".png", ".webp"],
    "Word": [".doc", ".docx"],
    "PPT": [".ppt", ".pptx"],
    "Archives": [".zip", ".rar", ".7z"],
    "Videos": [".mp4", ".mov", ".mkv"],
}
```

### 按日期归档和定时运行

如果文件很多，可以按修改时间创建 `2025-01` 这种月份文件夹。修改时间指文件最后一次被保存或改动的时间，可用 `file.stat().st_mtime` 获取，再用 `datetime` 格式化。

| 扩展需求 | 实现思路 | 适合难度 |
|---|---|---|
| 增加文件类型 | 在后缀映射字典中加入 `.docx`、`.pptx`、`.zip` 等 | 简单 |
| 按日期归档 | 根据修改时间创建 `YYYY-MM` 子文件夹 | 中等 |
| 定时运行 | Windows 用任务计划程序；macOS/Linux 用 cron（定时执行命令的系统工具） | 中等 |
| 递归整理子文件夹 | 使用 `rglob()` 遍历所有下级目录，递归指逐层进入子文件夹 | 中等偏难 |

规则不宜一开始就设计得太复杂。先满足“当前下载目录自动分类”这个需求，再逐步增加日期归档、子文件夹递归和定时运行，更容易排查错误。

## 常见问答(FAQ)

**Q:Python 自动整理文件夹需要安装第三方库吗？**

A:不需要安装第三方库。只整理文件名、后缀名和移动位置时，Python 标准库就够用，例如 pathlib 负责遍历目录，shutil 负责移动文件。新手只要安装好 Python 3，就可以直接运行脚本。

**Q:这个脚本能同时整理 Excel、PDF 和图片吗？**

A:可以同时整理 Excel、PDF 和图片。脚本通常按后缀名判断类型，例如 .xlsx、.xls 放入 Excel 文件夹，.pdf 放入 PDF 文件夹，.jpg、.png、.jpeg 放入图片文件夹。需要更多格式时继续添加后缀即可。

**Q:如果目标文件夹里已有同名文件怎么办？**

A:建议先做同名检查，避免直接覆盖原文件。常见做法是在新文件名后追加序号，例如 report.pdf 已存在，就改成 report_1.pdf 或 report_2.pdf 再移动。这样更适合新手批量整理。

**Q:能不能整理子文件夹里的文件？**

A:可以整理子文件夹里的文件，但遍历方式要改成递归。使用 pathlib 的 rglob('*') 可以扫描目标目录下所有层级文件，例如把 3 层子目录中的 PDF 统一移动到 PDF 文件夹。运行前要确认是否需要保留原目录结构。

**Q:新手运行前最需要注意什么？**

A:最需要先备份或用测试文件夹试跑。文件移动是实际改动目录结构的操作，建议先复制 10 个样例文件测试分类结果，确认 Excel、PDF、图片都进入正确文件夹后，再对正式目录运行。

<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "Article",
  "headline": "如何用 Python 自动整理文件夹里的 Excel、PDF 和图片文件？新手可直接套用的分类脚本",
  "keywords": [
    "Python 自动化",
    "Python 整理文件",
    "文件夹自动分类"
  ],
  "datePublished": "2026-05-31T00:14:06",
  "inLanguage": "zh-CN",
  "author": {
    "@type": "Organization",
    "name": "天机枢"
  },
  "publisher": {
    "@type": "Organization",
    "name": "天机枢"
  }
}
</script>
<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "FAQPage",
  "mainEntity": [
    {
      "@type": "Question",
      "name": "Python 自动整理文件夹需要安装第三方库吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "不需要安装第三方库。只整理文件名、后缀名和移动位置时，Python 标准库就够用，例如 pathlib 负责遍历目录，shutil 负责移动文件。新手只要安装好 Python 3，就可以直接运行脚本。"
      }
    },
    {
      "@type": "Question",
      "name": "这个脚本能同时整理 Excel、PDF 和图片吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "可以同时整理 Excel、PDF 和图片。脚本通常按后缀名判断类型，例如 .xlsx、.xls 放入 Excel 文件夹，.pdf 放入 PDF 文件夹，.jpg、.png、.jpeg 放入图片文件夹。需要更多格式时继续添加后缀即可。"
      }
    },
    {
      "@type": "Question",
      "name": "如果目标文件夹里已有同名文件怎么办？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "建议先做同名检查，避免直接覆盖原文件。常见做法是在新文件名后追加序号，例如 report.pdf 已存在，就改成 report_1.pdf 或 report_2.pdf 再移动。这样更适合新手批量整理。"
      }
    },
    {
      "@type": "Question",
      "name": "能不能整理子文件夹里的文件？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "可以整理子文件夹里的文件，但遍历方式要改成递归。使用 pathlib 的 rglob('*') 可以扫描目标目录下所有层级文件，例如把 3 层子目录中的 PDF 统一移动到 PDF 文件夹。运行前要确认是否需要保留原目录结构。"
      }
    },
    {
      "@type": "Question",
      "name": "新手运行前最需要注意什么？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "最需要先备份或用测试文件夹试跑。文件移动是实际改动目录结构的操作，建议先复制 10 个样例文件测试分类结果，确认 Excel、PDF、图片都进入正确文件夹后，再对正式目录运行。"
      }
    }
  ]
}
</script>
