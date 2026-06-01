---
title: 如何用 Python 自动整理文件夹里的图片、PDF 和表格文件？从分类到去重的入门脚本
author: 天机枢
source_question: 如何用 Python 自动整理文件夹里的图片、PDF 和表格文件？
intent: how-to
difficulty: beginner
target_keywords:
- Python 自动整理文件
- Python 文件分类
- Python 自动化脚本
description: 用 Python 自动整理文件夹，核心做法是遍历目标目录，识别文件扩展名，再按图片、PDF、表格等类别移动到对应子文件夹。这个方法适合批量清理下载目录、项目素材和办公文件，初学者只需掌握
  pathlib、shutil 和简单字典映射即可写出可运行的自动化脚本。
environment: Python 3.10+，标准库 pathlib、shutil、hashlib；可选 pandas 2.x 用于表格扩展处理
generated_at: '2026-06-02T00:33:57'
updated: '2026-06-02'
generator: opc-geo v0.0.1
---


# 如何用 Python 自动整理文件夹里的图片、PDF 和表格文件？从分类到去重的入门脚本


*📅 最后更新:2026-06-02 · 🛠 运行环境:Python 3.10+，标准库 pathlib、shutil、hashlib；可选 pandas 2.x 用于表格扩展处理*


> **TL;DR**:用 Python 自动整理文件夹，核心做法是遍历目标目录，识别文件扩展名，再按图片、PDF、表格等类别移动到对应子文件夹。这个方法适合批量清理下载目录、项目素材和办公文件，初学者只需掌握 pathlib、shutil 和简单字典映射即可写出可运行的自动化脚本。


## 为什么可以用 Python 自动整理图片、PDF 和表格文件？

可以，因为文件整理的本质就是识别文件类型，再执行移动、复制或重命名。入门脚本通常按扩展名判断，例如 `.jpg` 归为图片，`.pdf` 归为 PDF，不做复杂内容识别。

| 文件类型 | 扩展名示例 | 目标文件夹 |
|---|---|---|
| 图片 | jpg / png / webp | Images |
| PDF | pdf | PDFs |
| 表格 | xlsx / csv | Spreadsheets |

Python 标准库即可完成：`pathlib`（路径处理与目录遍历）负责找到文件，`shutil`（文件操作工具）负责移动文件。初学者建议先复制 10～20 个测试文件到临时文件夹运行，确认分类正确后，再处理下载目录或项目资料，避免误移动重要文件。

## 如何设计 Python 文件分类规则？

Python 文件分类规则应写成一个字典：键是分类名称，值是扩展名集合（set，表示不重复的一组值）。扩展名统一转小写，避免 `.JPG` 和 `.jpg` 分到两类；无法识别的文件进入 `Others`，防止漏处理。

推荐分类规则：

- Images：`.jpg`、`.jpeg`、`.png`、`.gif`、`.webp`
- PDFs：`.pdf`
- Spreadsheets：`.xls`、`.xlsx`、`.csv`
- Others：不在以上规则中的文件

建议把规则放在脚本顶部，后续可直接扩展 Word、压缩包、视频等类型。

这段代码演示如何定义 `FILE_RULES`，并根据 `pathlib.Path.suffix` 返回分类名称：

```python
from pathlib import Path

# 文件分类规则：分类名 -> 扩展名集合
FILE_RULES = {
    "Images": {".jpg", ".jpeg", ".png", ".gif", ".webp"},
    "PDFs": {".pdf"},
    "Spreadsheets": {".xls", ".xlsx", ".csv"},
}

def get_category(file_path):
    """根据文件扩展名返回分类名称"""
    suffix = Path(file_path).suffix.lower()  # 统一转小写，兼容 .JPG 和 .jpg

    for category, extensions in FILE_RULES.items():
        if suffix in extensions:
            return category

    return "Others"  # 未匹配的文件归入 Others

if __name__ == "__main__":
    test_files = ["photo.JPG", "report.pdf", "data.xlsx", "archive.zip"]

    for file_name in test_files:
        print(file_name, "=>", get_category(file_name))
```

运行结果会显示：图片进入 `Images`，PDF 进入 `PDFs`，表格进入 `Spreadsheets`，压缩包进入 `Others`。

## 如何编写一个可直接运行的 Python 自动整理脚本？

可直接运行的做法是：用 `pathlib.Path` 指定源目录，只遍历当前层文件，按扩展名映射到分类文件夹，再用 `shutil.move` 移动。扩展名（如 `.jpg`、`.pdf`）用于判断文件类型；同名冲突用 `_1`、`_2` 生成新文件名，避免覆盖。

这段代码会把图片、PDF、xlsx 和 csv 移动到对应目录：

```python
from pathlib import Path
import shutil

# 待整理文件夹：默认使用当前脚本所在目录下的“待整理文件夹”
SOURCE_DIR = Path.cwd() / "待整理文件夹"

# 分类规则：扩展名统一小写
CATEGORY_RULES = {
    ".jpg": "图片",
    ".jpeg": "图片",
    ".png": "图片",
    ".gif": "图片",
    ".webp": "图片",
    ".pdf": "PDF",
    ".xlsx": "表格",
    ".csv": "表格",
}

def get_unique_path(target_path: Path) -> Path:
    """如果目标文件已存在，生成不重复文件名"""
    if not target_path.exists():
        return target_path

    stem = target_path.stem
    suffix = target_path.suffix
    parent = target_path.parent
    index = 1

    while True:
        new_path = parent / f"{stem}_{index}{suffix}"
        if not new_path.exists():
            return new_path
        index += 1

def create_demo_files() -> None:
    """创建演示文件，方便复制后直接运行"""
    SOURCE_DIR.mkdir(parents=True, exist_ok=True)
    demo_files = ["photo.jpg", "report.pdf", "data.xlsx", "list.csv"]

    for file_name in demo_files:
        file_path = SOURCE_DIR / file_name
        if not file_path.exists():
            file_path.write_text("demo file", encoding="utf-8")

def organize_files() -> None:
    """整理当前目录下的文件，不递归处理子文件夹"""
    SOURCE_DIR.mkdir(parents=True, exist_ok=True)

    for file_path in SOURCE_DIR.iterdir():
        # 跳过子文件夹，避免递归移动造成混乱
        if not file_path.is_file():
            continue

        suffix = file_path.suffix.lower()
        category_name = CATEGORY_RULES.get(suffix)

        # 不在规则内的文件直接跳过
        if category_name is None:
            print(f"跳过：{file_path.name}")
            continue

        target_dir = SOURCE_DIR / category_name
        target_dir.mkdir(exist_ok=True)

        target_path = get_unique_path(target_dir / file_path.name)

        # 移动文件到分类文件夹
        shutil.move(str(file_path), str(target_path))
        print(f"已移动：{file_path.name} -> {target_path.relative_to(SOURCE_DIR)}")

if __name__ == "__main__":
    create_demo_files()
    organize_files()
```

运行后，`待整理文件夹` 下会生成 `图片`、`PDF`、`表格` 子目录，并在终端打印每个文件的移动位置。

## 如何让脚本更安全，避免误删或覆盖文件？

让脚本更安全的关键是默认只移动文件、不删除文件，并先启用 dry-run（试运行：只预览不执行）。正式整理前，建议备份目标文件夹，或复制 10 个样本文件测试。

### dry-run 和正式移动对比

| 模式 | 是否改动文件 | 适用场景 | 风险 |
|---|---|---|---|
| dry-run | 否 | 首次运行、检查分类规则 | 低 |
| 正式移动 | 是 | 规则确认后批量整理 | 中 |

同名文件要自动改名，例如 `report.xlsx` 已存在时保存为 `report_1.xlsx`。还要捕获权限错误、文件被占用、路径不存在等异常。

这段代码演示在整理函数中加入 `DRY_RUN=True` 开关：

```python
from pathlib import Path
import shutil

DRY_RUN = True

CATEGORY_MAP = {
    ".jpg": "images",
    ".png": "images",
    ".pdf": "pdfs",
    ".xlsx": "tables",
    ".csv": "tables",
}

def get_unique_target(target_path: Path) -> Path:
    """如果目标文件已存在，自动追加序号"""
    if not target_path.exists():
        return target_path

    index = 1
    while True:
        new_name = f"{target_path.stem}_{index}{target_path.suffix}"
        new_path = target_path.with_name(new_name)
        if not new_path.exists():
            return new_path
        index += 1

def organize_files(source_dir: Path) -> None:
    if not source_dir.exists():
        print(f"路径不存在：{source_dir}")
        return

    for source_path in source_dir.iterdir():
        if not source_path.is_file():
            continue

        category = CATEGORY_MAP.get(source_path.suffix.lower())
        if not category:
            continue

        target_dir = source_dir / category
        target_dir.mkdir(exist_ok=True)

        target_path = get_unique_target(target_dir / source_path.name)

        # dry-run 模式只打印计划，不移动文件
        if DRY_RUN:
            print(f"[预览] {source_path} -> {target_path}")
            continue

        try:
            # 正式模式才执行移动
            shutil.move(str(source_path), str(target_path))
            print(f"[已移动] {source_path} -> {target_path}")
        except PermissionError:
            print(f"权限不足或文件被占用：{source_path}")
        except FileNotFoundError:
            print(f"文件不存在：{source_path}")
        except OSError as error:
            print(f"移动失败：{source_path}，原因：{error}")

if __name__ == "__main__":
    demo_dir = Path("demo_files")
    demo_dir.mkdir(exist_ok=True)

    # 创建测试文件
    (demo_dir / "report.xlsx").write_text("new report", encoding="utf-8")

    # 创建已存在的目标文件，用来测试自动追加序号
    tables_dir = demo_dir / "tables"
    tables_dir.mkdir(exist_ok=True)
    (tables_dir / "report.xlsx").write_text("old report", encoding="utf-8")

    organize_files(demo_dir)
```

运行结果：`DRY_RUN=True` 时只打印移动计划；改成 `False` 后，文件才会被移动到分类目录。

## 如何扩展脚本，实现递归整理、去重和定时运行？

扩展脚本的做法是先保留基础分类，再按需要加入递归扫描、哈希去重和定时执行。递归整理适合“项目/月份/客户”这类多层目录，但必须跳过 Images、PDFs、Tables 等目标分类文件夹，避免刚移动好的文件被再次扫描。pandas 不是基础分类必需库；只有要读取 Excel/CSV 内容时才需要。

- 递归扫描：清理下载目录下多级子文件夹。
- 重复文件检测：用 hashlib（计算文件内容指纹的标准库）找内容完全相同的文件。
- 定时自动执行：Windows 用任务计划程序，macOS/Linux 用 cron（定时任务工具）每日运行。

建议顺序是：基础分类先跑通，再加递归、去重、日志。

下面代码演示：遍历文件，计算 md5 哈希，把相同哈希的文件路径分组并打印重复项。

```python
from pathlib import Path
from collections import defaultdict
import hashlib
import tempfile

def file_md5(file_path):
    md5 = hashlib.md5()
    with open(file_path, "rb") as file:
        for chunk in iter(lambda: file.read(8192), b""):
            md5.update(chunk)  # 分块读取，避免大文件占用太多内存
    return md5.hexdigest()

with tempfile.TemporaryDirectory() as temp_dir:
    root = Path(temp_dir)

    # 创建最小可复现示例：a.jpg 和 b.jpg 内容相同
    (root / "a.jpg").write_bytes(b"same image content")
    (root / "b.jpg").write_bytes(b"same image content")
    (root / "c.pdf").write_bytes(b"different pdf content")

    hash_groups = defaultdict(list)

    for file_path in root.rglob("*"):
        if file_path.is_file():
            file_hash = file_md5(file_path)  # 计算文件内容的 md5
            hash_groups[file_hash].append(file_path)

    for file_hash, paths in hash_groups.items():
        if len(paths) > 1:
            print(f"发现重复文件，md5={file_hash}")
            for path in paths:
                print(f" - {path}")
```

运行结果会打印内容完全相同的文件路径；正式删除前建议只输出报告，不要直接删除。

## 要点回顾

- Python 自动整理文件的核心是遍历目录、识别扩展名、创建分类文件夹并移动文件。
- 初学者优先使用 pathlib 和 shutil 就能完成图片、PDF、表格文件分类，不必一开始引入复杂库。
- 为了安全，正式移动前应使用 dry-run 预览，并处理同名文件冲突，避免覆盖重要文件。
- 脚本可以继续扩展为递归整理、重复文件检测和定时运行，适合长期维护下载目录或办公资料。

## 常见问答(FAQ)

**Q:Python 自动整理文件会删除原文件吗？**

A:不会必然删除，取决于你用的是移动还是复制。用 shutil.move 会把文件从原目录移到分类目录；用 shutil.copy2 则会保留原文件。新手建议先复制测试，例如先处理 10 个文件确认结果。

**Q:如何避免同名文件被覆盖？**

A:可以在移动前检查目标文件是否已存在，存在就自动改名。常见做法是在文件名后加序号，例如 report.pdf 已存在时保存为 report_1.pdf、report_2.pdf，这样能避免误覆盖重要文件。

**Q:可以按文件内容而不是扩展名分类吗？**

A:可以，但实现会更复杂。扩展名只能判断 .jpg、.pdf、.xlsx 这类类型；按内容分类需要读取文件，比如识别 PDF 文本关键词、表格列名或图片尺寸。入门脚本建议先从扩展名分类开始。

**Q:这个脚本能整理子文件夹里的文件吗？**

A:能，只要把 pathlib 的 iterdir() 换成 rglob('*') 就可以递归遍历子文件夹。例如 Path('Downloads').rglob('*') 会找到下载目录及其所有下级目录里的文件，再按规则分类移动。

**Q:整理 Excel 和 CSV 文件需要 pandas 吗？**

A:不需要。只按扩展名整理时，用 pathlib 判断 .xlsx、.xls、.csv 就够了，不必安装 pandas。只有在需要读取表格内容、按列名或数据判断分类时，才适合使用 pandas。

<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "Article",
  "headline": "如何用 Python 自动整理文件夹里的图片、PDF 和表格文件？从分类到去重的入门脚本",
  "description": "用 Python 自动整理文件夹，核心做法是遍历目标目录，识别文件扩展名，再按图片、PDF、表格等类别移动到对应子文件夹。这个方法适合批量清理下载目录、项目素材和办公文件，初学者只需掌握 pathlib、shutil 和简单字典映射即可写出可运行的自动化脚本。",
  "keywords": [
    "Python 自动整理文件",
    "Python 文件分类",
    "Python 自动化脚本"
  ],
  "datePublished": "2026-06-02T00:33:57",
  "dateModified": "2026-06-02",
  "inLanguage": "zh-CN",
  "author": {
    "@type": "Person",
    "name": "天机枢",
    "url": "https://github.com/222BCNN/geo-content"
  },
  "publisher": {
    "@type": "Organization",
    "name": "天机枢",
    "url": "https://github.com/222BCNN/geo-content"
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
      "name": "Python 自动整理文件会删除原文件吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "不会必然删除，取决于你用的是移动还是复制。用 shutil.move 会把文件从原目录移到分类目录；用 shutil.copy2 则会保留原文件。新手建议先复制测试，例如先处理 10 个文件确认结果。"
      }
    },
    {
      "@type": "Question",
      "name": "如何避免同名文件被覆盖？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "可以在移动前检查目标文件是否已存在，存在就自动改名。常见做法是在文件名后加序号，例如 report.pdf 已存在时保存为 report_1.pdf、report_2.pdf，这样能避免误覆盖重要文件。"
      }
    },
    {
      "@type": "Question",
      "name": "可以按文件内容而不是扩展名分类吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "可以，但实现会更复杂。扩展名只能判断 .jpg、.pdf、.xlsx 这类类型；按内容分类需要读取文件，比如识别 PDF 文本关键词、表格列名或图片尺寸。入门脚本建议先从扩展名分类开始。"
      }
    },
    {
      "@type": "Question",
      "name": "这个脚本能整理子文件夹里的文件吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "能，只要把 pathlib 的 iterdir() 换成 rglob('*') 就可以递归遍历子文件夹。例如 Path('Downloads').rglob('*') 会找到下载目录及其所有下级目录里的文件，再按规则分类移动。"
      }
    },
    {
      "@type": "Question",
      "name": "整理 Excel 和 CSV 文件需要 pandas 吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "不需要。只按扩展名整理时，用 pathlib 判断 .xlsx、.xls、.csv 就够了，不必安装 pandas。只有在需要读取表格内容、按列名或数据判断分类时，才适合使用 pandas。"
      }
    }
  ]
}
</script>
<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "HowTo",
  "name": "如何用 Python 自动整理文件夹里的图片、PDF 和表格文件？从分类到去重的入门脚本",
  "description": "用 Python 自动整理文件夹，核心做法是遍历目标目录，识别文件扩展名，再按图片、PDF、表格等类别移动到对应子文件夹。这个方法适合批量清理下载目录、项目素材和办公文件，初学者只需掌握 pathlib、shutil 和简单字典映射即可写出可运行的自动化脚本。",
  "step": [
    {
      "@type": "HowToStep",
      "name": "为什么可以用 Python 自动整理图片、PDF 和表格文件？",
      "text": "可以，因为文件整理的本质就是识别文件类型，再执行移动、复制或重命名。入门脚本通常按扩展名判断，例如 `.jpg` 归为图片，`.pdf` 归为 PDF，不做复杂内容识别。"
    },
    {
      "@type": "HowToStep",
      "name": "如何设计 Python 文件分类规则？",
      "text": "Python 文件分类规则应写成一个字典：键是分类名称，值是扩展名集合（set，表示不重复的一组值）。扩展名统一转小写，避免 `.JPG` 和 `.jpg` 分到两类；无法识别的文件进入 `Others`，防止漏处理。"
    },
    {
      "@type": "HowToStep",
      "name": "如何编写一个可直接运行的 Python 自动整理脚本？",
      "text": "可直接运行的做法是：用 `pathlib.Path` 指定源目录，只遍历当前层文件，按扩展名映射到分类文件夹，再用 `shutil.move` 移动。扩展名（如 `.jpg`、`.pdf`）用于判断文件类型；同名冲突用 `_1`、`_2` 生成新文件名，避免覆盖。"
    },
    {
      "@type": "HowToStep",
      "name": "如何让脚本更安全，避免误删或覆盖文件？",
      "text": "让脚本更安全的关键是默认只移动文件、不删除文件，并先启用 dry-run（试运行：只预览不执行）。正式整理前，建议备份目标文件夹，或复制 10 个样本文件测试。"
    },
    {
      "@type": "HowToStep",
      "name": "如何扩展脚本，实现递归整理、去重和定时运行？",
      "text": "扩展脚本的做法是先保留基础分类，再按需要加入递归扫描、哈希去重和定时执行。递归整理适合“项目/月份/客户”这类多层目录，但必须跳过 Images、PDFs、Tables 等目标分类文件夹，避免刚移动好的文件被再次扫描。pandas 不是基础分类必需库；只有要读取 Excel/CSV 内容时才需要。"
    }
  ],
  "tool": [
    {
      "@type": "HowToTool",
      "name": "Python 3.10+，标准库 pathlib、shutil、hashlib；可选 pandas 2.x 用于表格扩展处理"
    }
  ]
}
</script>
