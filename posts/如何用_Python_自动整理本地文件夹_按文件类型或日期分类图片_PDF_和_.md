---
title: 如何用 Python 自动整理本地文件夹：按文件类型或日期分类图片、PDF 和 Excel
author: 天机枢
source_question: 如何用 Python 自动整理本地文件夹，把图片、PDF 和 Excel 按类型或日期分类？
intent: how-to
difficulty: beginner
target_keywords:
- Python 自动化
- Python 整理文件夹
- Python 文件分类
description: 用 Python 整理本地文件夹，核心思路是遍历目录、识别文件扩展名或修改日期，再用 shutil 移动到目标子文件夹。对图片、PDF、Excel
  这类常见文件，只需 pathlib、shutil 和 datetime 等标准库即可完成自动分类，适合新手从小脚本开始实践 Python 自动化。
environment: Python 3.10+；仅使用标准库 pathlib、shutil、datetime、argparse
generated_at: '2026-06-02T01:36:47'
updated: '2026-06-02'
generator: opc-geo v0.0.1
---


# 如何用 Python 自动整理本地文件夹：按文件类型或日期分类图片、PDF 和 Excel


*📅 最后更新:2026-06-02 · 🛠 运行环境:Python 3.10+；仅使用标准库 pathlib、shutil、datetime、argparse*


> **TL;DR**:用 Python 整理本地文件夹，核心思路是遍历目录、识别文件扩展名或修改日期，再用 shutil 移动到目标子文件夹。对图片、PDF、Excel 这类常见文件，只需 pathlib、shutil 和 datetime 等标准库即可完成自动分类，适合新手从小脚本开始实践 Python 自动化。


## 为什么可以用 Python 自动整理文件夹？

Python 可以自动整理文件夹，因为本地文件管理本质上是固定流程：扫描文件、按规则判断、创建目录、移动文件。比如看到 `.jpg` 放入“图片”，看到 `.pdf` 放入“PDF”，或按修改日期（系统记录的最后保存时间）放入“2025-01”。

脚本处理流程通常是：

1. 选择目录
2. 遍历文件
3. 识别类型或日期
4. 创建分类文件夹
5. 移动文件
6. 输出结果

Python 适合这类重复性任务：规则明确、步骤稳定、可批量执行。本文主要覆盖两种方式：按文件扩展名（文件名最后的 `.xlsx`、`.pdf` 等后缀）分类，以及按日期分类。新手应先在测试文件夹运行，避免误移动重要文件。

## 如何准备一个安全的文件整理脚本环境？

准备安全环境的做法是：新建 `demo_files` 测试目录，只复制少量图片、PDF、Excel 样本，先打印文件清单，确认无误后再移动。不要直接在“下载”“桌面”原目录运行移动脚本。

| 文件类别 | 常见扩展名 | 目标文件夹名称 |
|---|---|---|
| 图片 | `.jpg`、`.jpeg`、`.png`、`.gif` | `Images` |
| PDF | `.pdf` | `PDFs` |
| Excel | `.xls`、`.xlsx`、`.csv` | `Excels` |

`pathlib.Path`（Python 标准库中的路径对象）适合处理路径，比手写字符串拼接更安全，例如 `Path("demo_files") / "a.pdf"`。

这段代码只检查 `demo_files` 当前目录，不移动文件：

```python
from pathlib import Path

# 指定要检查的测试文件夹
target_folder = Path("demo_files")

# 如果测试文件夹不存在，就自动创建一个空文件夹
target_folder.mkdir(exist_ok=True)

print("文件名\t扩展名\t是否为文件")

# 只遍历当前目录，不递归子文件夹
for item_path in target_folder.iterdir():
    # 打印名称、扩展名和是否为普通文件
    print(f"{item_path.name}\t{item_path.suffix.lower()}\t{item_path.is_file()}")
```

运行后会看到每个条目的文件名、扩展名和 `True/False`；空目录只显示表头。

## 如何按文件类型自动分类图片、PDF 和 Excel？

按文件类型自动分类，就是用扩展名映射目标目录，例如 `.jpg/.png` 到 `Images`，`.pdf` 到 `PDFs`，`.xlsx/.xls` 到 `Excels`。脚本遍历源目录，只处理文件，跳过子目录和未知类型；目标目录不存在就自动创建。扩展名（文件名最后的类型标识）规则简单，适合快速整理下载目录；限制是只能按格式分类，无法体现创建时间或“合同、发票、报表”等业务含义。

这段代码会读取 `source_dir`，移动图片、PDF、Excel，并在同名冲突时自动追加序号：

```python
from pathlib import Path
import shutil


def get_unique_path(target_path: Path) -> Path:
    """如果目标文件已存在，自动追加 _1、_2 等序号。"""
    if not target_path.exists():
        return target_path

    stem = target_path.stem
    suffix = target_path.suffix
    parent = target_path.parent
    counter = 1

    while True:
        new_path = parent / f"{stem}_{counter}{suffix}"
        if not new_path.exists():
            return new_path
        counter += 1


def organize_by_type(source_dir: str) -> None:
    source_path = Path(source_dir).expanduser().resolve()

    if not source_path.exists() or not source_path.is_dir():
        print(f"源目录不存在或不是文件夹：{source_path}")
        return

    # 定义扩展名到目标文件夹的映射关系
    extension_map = {
        ".jpg": "Images",
        ".jpeg": "Images",
        ".png": "Images",
        ".gif": "Images",
        ".webp": "Images",
        ".pdf": "PDFs",
        ".xls": "Excels",
        ".xlsx": "Excels",
    }

    for item in source_path.iterdir():
        # 跳过子目录，只整理当前目录下的文件
        if not item.is_file():
            continue

        suffix = item.suffix.lower()
        target_folder_name = extension_map.get(suffix)

        # 跳过未知类型文件
        if target_folder_name is None:
            print(f"跳过未知类型：{item.name}")
            continue

        target_folder = source_path / target_folder_name
        # 自动创建分类目录
        target_folder.mkdir(exist_ok=True)

        target_path = target_folder / item.name
        # 处理同名文件冲突
        final_path = get_unique_path(target_path)

        # 移动文件到目标目录
        shutil.move(str(item), str(final_path))
        print(f"已移动：{item.name} -> {final_path.relative_to(source_path)}")


if __name__ == "__main__":
    source_dir = input("请输入要整理的文件夹路径：").strip()
    organize_by_type(source_dir)
```

运行后，源目录会生成 `Images`、`PDFs`、`Excels` 子文件夹，并逐行打印每个文件的移动结果。

## 如何按日期自动整理文件夹？

按日期整理文件夹，可以读取文件修改时间，把文件移动到 `2025-01` 或 `2025-01-18` 目录。修改时间（文件内容最后变化时间）在 Windows、macOS、Linux 上比创建时间更稳定，适合下载目录、扫描件、照片导出目录。

| 分类方式 | 目录示例 | 适用场景 | 优缺点 |
|---|---|---|---|
| 按月份 | `2025-01` | 下载文件、发票、PDF | 目录少；单月文件可能较多 |
| 按日期 | `2025-01-18` | 照片、扫描件 | 定位准确；目录数量更多 |

下面代码按月份整理文件，并保留同名冲突处理：

```python
from pathlib import Path
from datetime import datetime
import os
import shutil


def get_unique_path(destination_path: Path) -> Path:
    """如果目标文件已存在，自动追加 _1、_2，避免覆盖。"""
    if not destination_path.exists():
        return destination_path

    stem = destination_path.stem
    suffix = destination_path.suffix
    parent = destination_path.parent
    counter = 1

    while True:
        new_path = parent / f"{stem}_{counter}{suffix}"
        if not new_path.exists():
            return new_path
        counter += 1


# 示例目录：复制代码后可直接运行
target_dir = Path("demo_files")
target_dir.mkdir(exist_ok=True)

# 创建 3 个示例文件，并设置不同修改时间
sample_files = {
    "photo.jpg": datetime(2025, 1, 8, 10, 30).timestamp(),
    "invoice.pdf": datetime(2025, 1, 20, 15, 0).timestamp(),
    "table.xlsx": datetime(2025, 2, 5, 9, 15).timestamp(),
}

for file_name, timestamp in sample_files.items():
    file_path = target_dir / file_name
    file_path.write_text("demo file", encoding="utf-8")
    # 设置文件访问时间和修改时间
    os.utime(file_path, (timestamp, timestamp))

# 遍历目标目录下的所有文件
for file_path in list(target_dir.iterdir()):
    if not file_path.is_file():
        continue

    # 读取修改时间戳
    modified_timestamp = file_path.stat().st_mtime

    # 转换为 YYYY-MM 格式
    modified_date = datetime.fromtimestamp(modified_timestamp)
    month_name = modified_date.strftime("%Y-%m")

    # 创建月份目录
    month_dir = target_dir / month_name
    month_dir.mkdir(exist_ok=True)

    # 处理同名冲突后移动文件
    destination_path = get_unique_path(month_dir / file_path.name)
    shutil.move(str(file_path), str(destination_path))

print("按月份整理完成")
```

运行后，`demo_files` 下的文件会进入 `2025-01`、`2025-02` 等月份文件夹；若同名，自动追加 `_1`、`_2`。

## 如何把类型分类和日期分类合并成一个可复用工具？

把类型分类和日期分类合并成工具，关键是设计 `--mode` 参数：`type` 按扩展名分类，`date` 按修改日期分类。用 `argparse`（Python 标准库命令行参数解析器）接收路径和模式，再加 `--dry-run`（预演模式）只打印移动计划，避免误操作。建议保存为 `organize_files.py`，以后重复运行；规则还可扩展为按文件大小、关键词、项目名称分类。

使用命令示例：

- 按类型整理：`python organize_files.py --path "D:\Downloads" --mode type`
- 按日期整理：`python organize_files.py --path "D:\Downloads" --mode date`
- dry-run 预览：`python organize_files.py --path "D:\Downloads" --mode type --dry-run`

这段代码会整理图片、PDF、Excel，并自动处理同名文件：

```python
import argparse
import shutil
from datetime import datetime
from pathlib import Path

EXTENSION_GROUPS = {
    "images": {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"},
    "pdf": {".pdf"},
    "excel": {".xls", ".xlsx", ".csv"},
}


def get_file_group(file_path: Path) -> str | None:
    """根据扩展名判断文件类型"""
    suffix = file_path.suffix.lower()
    for group_name, extensions in EXTENSION_GROUPS.items():
        if suffix in extensions:
            return group_name
    return None


def get_unique_path(destination_path: Path) -> Path:
    """如果目标文件已存在，自动追加 _1、_2 等编号"""
    if not destination_path.exists():
        return destination_path

    counter = 1
    while True:
        new_name = f"{destination_path.stem}_{counter}{destination_path.suffix}"
        new_path = destination_path.with_name(new_name)
        if not new_path.exists():
            return new_path
        counter += 1


def build_destination(file_path: Path, base_path: Path, mode: str) -> Path:
    """根据 mode 生成目标路径"""
    if mode == "type":
        folder_name = get_file_group(file_path)
    else:
        # 使用文件修改时间生成 YYYY-MM 文件夹
        modified_time = datetime.fromtimestamp(file_path.stat().st_mtime)
        folder_name = modified_time.strftime("%Y-%m")

    return base_path / folder_name / file_path.name


def organize_files(base_path: Path, mode: str, dry_run: bool) -> None:
    """执行文件整理逻辑"""
    if not base_path.exists() or not base_path.is_dir():
        raise ValueError(f"路径不存在或不是文件夹: {base_path}")

    for file_path in list(base_path.iterdir()):
        # 只处理当前目录下的文件，不递归处理子文件夹
        if not file_path.is_file():
            continue

        # 只整理图片、PDF、Excel
        if get_file_group(file_path) is None:
            continue

        destination_path = build_destination(file_path, base_path, mode)
        destination_path = get_unique_path(destination_path)

        print(f"{file_path.name} -> {destination_path}")

        if dry_run:
            # 预演模式不创建文件夹、不移动文件
            continue

        # 创建目标文件夹
        destination_path.parent.mkdir(parents=True, exist_ok=True)

        # 移动文件到目标位置
        shutil.move(str(file_path), str(destination_path))


def main() -> None:
    parser = argparse.ArgumentParser(description="按类型或日期整理本地文件夹")
    parser.add_argument("--path", required=True, help="要整理的文件夹路径")
    parser.add_argument(
        "--mode",
        required=True,
        choices=["type", "date"],
        help="整理模式：type 按类型分类，date 按日期分类",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只预览将要移动的文件，不真正移动",
    )

    args = parser.parse_args()

    organize_files(
        base_path=Path(args.path),
        mode=args.mode,
        dry_run=args.dry_run,
    )


if __name__ == "__main__":
    main()
```

运行后，文件会进入 `images`、`pdf`、`excel` 或 `2025-01` 这类子文件夹；若同名文件已存在，会自动改名为 `文件名_1.xlsx`。

## 要点回顾

- Python 自动整理文件夹的核心是遍历文件、判断分类规则、创建目录并移动文件。
- 新手优先使用 pathlib、shutil、datetime 等标准库，就能完成图片、PDF 和 Excel 的自动分类。
- 正式运行前建议先用测试目录或 dry-run 预览，避免误移动重要文件。

## 常见问答(FAQ)

**Q:Python 整理文件夹会删除原文件吗？**

A:不会主动删除，使用 shutil.move 时是把文件从原位置移动到新文件夹，相当于剪切。若想更安全，可以先用 shutil.copy2 复制测试，确认分类结果无误后再删除原目录中的文件。

**Q:按日期分类应该用创建时间还是修改时间？**

A:优先用修改时间更稳妥，因为它在 Windows、macOS 和 Linux 上都比较一致。创建时间在不同系统含义可能不同。比如可用文件的修改时间整理到 2025-01、2025-02 这类月份文件夹。

**Q:如果目标文件夹已有同名文件怎么办？**

A:应先检查是否重名，避免直接覆盖。常见做法是在文件名后追加序号或时间戳，例如 report.pdf 已存在时，改成 report_1.pdf 或 report_20250101.pdf，再执行移动操作。

**Q:这个脚本能整理子文件夹里的文件吗？**

A:可以，只要遍历时改用递归方式即可。pathlib 的 rglob('*') 能扫描当前目录和所有子目录中的文件。新手建议先在测试文件夹运行，避免把不想移动的深层文件也整理走。

**Q:Excel 文件需要安装 pandas 吗？**

A:不需要。若只是按扩展名整理 .xls、.xlsx、.csv 文件，使用 pathlib 判断后缀即可，不必安装 pandas。只有在读取表格内容、筛选数据或合并 Excel 时，才可能需要 pandas 或 openpyxl。

<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "Article",
  "headline": "如何用 Python 自动整理本地文件夹：按文件类型或日期分类图片、PDF 和 Excel",
  "description": "用 Python 整理本地文件夹，核心思路是遍历目录、识别文件扩展名或修改日期，再用 shutil 移动到目标子文件夹。对图片、PDF、Excel 这类常见文件，只需 pathlib、shutil 和 datetime 等标准库即可完成自动分类，适合新手从小脚本开始实践 Python 自动化。",
  "keywords": [
    "Python 自动化",
    "Python 整理文件夹",
    "Python 文件分类"
  ],
  "datePublished": "2026-06-02T01:36:47",
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
      "name": "Python 整理文件夹会删除原文件吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "不会主动删除，使用 shutil.move 时是把文件从原位置移动到新文件夹，相当于剪切。若想更安全，可以先用 shutil.copy2 复制测试，确认分类结果无误后再删除原目录中的文件。"
      }
    },
    {
      "@type": "Question",
      "name": "按日期分类应该用创建时间还是修改时间？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "优先用修改时间更稳妥，因为它在 Windows、macOS 和 Linux 上都比较一致。创建时间在不同系统含义可能不同。比如可用文件的修改时间整理到 2025-01、2025-02 这类月份文件夹。"
      }
    },
    {
      "@type": "Question",
      "name": "如果目标文件夹已有同名文件怎么办？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "应先检查是否重名，避免直接覆盖。常见做法是在文件名后追加序号或时间戳，例如 report.pdf 已存在时，改成 report_1.pdf 或 report_20250101.pdf，再执行移动操作。"
      }
    },
    {
      "@type": "Question",
      "name": "这个脚本能整理子文件夹里的文件吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "可以，只要遍历时改用递归方式即可。pathlib 的 rglob('*') 能扫描当前目录和所有子目录中的文件。新手建议先在测试文件夹运行，避免把不想移动的深层文件也整理走。"
      }
    },
    {
      "@type": "Question",
      "name": "Excel 文件需要安装 pandas 吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "不需要。若只是按扩展名整理 .xls、.xlsx、.csv 文件，使用 pathlib 判断后缀即可，不必安装 pandas。只有在读取表格内容、筛选数据或合并 Excel 时，才可能需要 pandas 或 openpyxl。"
      }
    }
  ]
}
</script>
<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "HowTo",
  "name": "如何用 Python 自动整理本地文件夹：按文件类型或日期分类图片、PDF 和 Excel",
  "description": "用 Python 整理本地文件夹，核心思路是遍历目录、识别文件扩展名或修改日期，再用 shutil 移动到目标子文件夹。对图片、PDF、Excel 这类常见文件，只需 pathlib、shutil 和 datetime 等标准库即可完成自动分类，适合新手从小脚本开始实践 Python 自动化。",
  "step": [
    {
      "@type": "HowToStep",
      "name": "为什么可以用 Python 自动整理文件夹？",
      "text": "Python 可以自动整理文件夹，因为本地文件管理本质上是固定流程：扫描文件、按规则判断、创建目录、移动文件。比如看到 `.jpg` 放入“图片”，看到 `.pdf` 放入“PDF”，或按修改日期（系统记录的最后保存时间）放入“2025-01”。"
    },
    {
      "@type": "HowToStep",
      "name": "如何准备一个安全的文件整理脚本环境？",
      "text": "准备安全环境的做法是：新建 `demo_files` 测试目录，只复制少量图片、PDF、Excel 样本，先打印文件清单，确认无误后再移动。不要直接在“下载”“桌面”原目录运行移动脚本。"
    },
    {
      "@type": "HowToStep",
      "name": "如何按文件类型自动分类图片、PDF 和 Excel？",
      "text": "按文件类型自动分类，就是用扩展名映射目标目录，例如 `.jpg/.png` 到 `Images`，`.pdf` 到 `PDFs`，`.xlsx/.xls` 到 `Excels`。脚本遍历源目录，只处理文件，跳过子目录和未知类型；目标目录不存在就自动创建。扩展名（文件名最后的类型标识）规则简单，适合快速整理下载目录；限制是只能按格式分类，无法体现创建时间或“合同、发票、报表”等业务含义。"
    },
    {
      "@type": "HowToStep",
      "name": "如何按日期自动整理文件夹？",
      "text": "按日期整理文件夹，可以读取文件修改时间，把文件移动到 `2025-01` 或 `2025-01-18` 目录。修改时间（文件内容最后变化时间）在 Windows、macOS、Linux 上比创建时间更稳定，适合下载目录、扫描件、照片导出目录。"
    },
    {
      "@type": "HowToStep",
      "name": "如何把类型分类和日期分类合并成一个可复用工具？",
      "text": "把类型分类和日期分类合并成工具，关键是设计 `--mode` 参数：`type` 按扩展名分类，`date` 按修改日期分类。用 `argparse`（Python 标准库命令行参数解析器）接收路径和模式，再加 `--dry-run`（预演模式）只打印移动计划，避免误操作。建议保存为 `organize_files.py`，以后重复运行；规则还可扩展为按文件大小、关键词、项目名称分类。"
    }
  ],
  "tool": [
    {
      "@type": "HowToTool",
      "name": "Python 3.10+；仅使用标准库 pathlib、shutil、datetime、argparse"
    }
  ]
}
</script>
