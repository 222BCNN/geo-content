---
title: 如何用 Python 自动批量处理 Excel 表格并生成汇总报表？新手完整流程
author: 天机枢
source_question: 如何用 Python 自动批量处理 Excel 表格并生成汇总报表？
intent: how-to
difficulty: beginner
target_keywords:
- Python 自动化
- Python 处理 Excel
- 批量生成报表
description: 用 Python 批量处理 Excel 的核心流程是：读取多个文件、清洗和合并数据、按规则统计汇总，最后导出新的报表。新手建议优先使用 pandas
  处理数据，用 pathlib 扫描文件，用 openpyxl 辅助设置 Excel 格式，就能把重复的人工统计工作自动化。
environment: Python 3.12、pandas 2.x、openpyxl 3.x、Windows/macOS/Linux 均可
generated_at: '2026-06-04T09:34:34'
updated: '2026-06-04'
generator: opc-geo v0.0.1
---


# 如何用 Python 自动批量处理 Excel 表格并生成汇总报表？新手完整流程


*📅 最后更新:2026-06-04 · 🛠 运行环境:Python 3.12、pandas 2.x、openpyxl 3.x、Windows/macOS/Linux 均可*


> **TL;DR**:用 Python 批量处理 Excel 的核心流程是：读取多个文件、清洗和合并数据、按规则统计汇总，最后导出新的报表。新手建议优先使用 pandas 处理数据，用 pathlib 扫描文件，用 openpyxl 辅助设置 Excel 格式，就能把重复的人工统计工作自动化。


## 为什么适合用 Python 自动化处理 Excel 报表？

Python 适合自动化处理 Excel 报表，因为它能把“逐个打开文件、复制粘贴、手工统计”变成可重复运行的脚本。人工处理销售日报、库存表、考勤表、财务明细、项目进度表时，常见问题是耗时、漏算、列名不一致、日期格式混乱。

新手建议优先用 pandas（Python 的表格数据处理库），先把 Excel 当成数据表处理，而不是一开始操作复杂的单元格对象。

适合自动化的 Excel 任务包括：

- 批量合并多个 Excel 文件
- 清洗字段，如去空格、统一日期
- 按部门、月份、产品分组统计
- 生成汇总表和明细表
- 导出新的 Excel 报表

## 如何准备 Python 处理 Excel 的运行环境？

准备环境只需安装 Python 3.10+、pandas 和 openpyxl，并建立固定目录。pandas（数据处理库）负责合并统计，openpyxl（xlsx 读写库）负责 Excel 文件读写和格式。

| 工具名称 | 作用 | 安装命令 |
|---|---|---|
| pandas | 表格清洗、汇总 | `pip install pandas` |
| openpyxl | 读写 `.xlsx` | `pip install openpyxl` |
| pathlib | 扫描文件路径 | Python 内置，无需安装 |

推荐目录：

```text
excel_report/
├─ input/      # 原始 Excel
├─ output/     # 汇总报表
└─ main.py     # 自动化脚本
```

Excel 建议统一为 `.xlsx`，且每个文件表头一致，例如都包含“日期、部门、金额”。常见问题：先升级 pip（Python 包安装工具），优先使用虚拟环境（独立依赖目录），路径避免空格、中文括号、`#` 等特殊字符。

这段代码会自动安装依赖，并检查版本：

```python
import sys
import subprocess
import importlib.metadata

# 使用当前 Python 环境安装 pandas 和 openpyxl
subprocess.check_call([
    sys.executable, "-m", "pip", "install", "--upgrade", "pandas", "openpyxl"
])

# 检查依赖是否安装成功
pandas_version = importlib.metadata.version("pandas")
openpyxl_version = importlib.metadata.version("openpyxl")

print(f"pandas 版本: {pandas_version}")
print(f"openpyxl 版本: {openpyxl_version}")
```

运行后看到两个版本号，即表示环境可用。

## 如何用 Python 批量读取多个 Excel 文件？

用 Python 批量读取多个 Excel 文件，可以用 `pathlib` 扫描 `input` 文件夹中的所有 `.xlsx`，再用 `pandas.read_excel` 逐个读成 DataFrame（pandas 的二维表格数据结构）并合并。

批量读取流程：

1. 扫描文件：获取 `input/*.xlsx` 文件列表。
2. 循环读取：对每个文件执行 `pd.read_excel()`。
3. 追加来源字段：新增 `source_file` 列记录文件名。
4. 合并总表：用 `pd.concat()` 合成一个总表。

这段代码会读取 `input` 文件夹的所有 `.xlsx`，合并后打印前 5 行：

```python
from pathlib import Path
import pandas as pd

input_dir = Path("input")
input_dir.mkdir(exist_ok=True)

# 如果 input 文件夹为空，创建两个示例 Excel，保证代码可直接运行
if not list(input_dir.glob("*.xlsx")):
    pd.DataFrame({"部门": ["销售部", "技术部"], "金额": [1200, 800]}).to_excel(input_dir / "一月.xlsx", index=False)
    pd.DataFrame({"部门": ["销售部", "行政部"], "金额": [1500, 600]}).to_excel(input_dir / "二月.xlsx", index=False)

dataframes = []

# 扫描 input 文件夹中的所有 .xlsx 文件
for workbook_path in input_dir.glob("*.xlsx"):
    # 读取单个 Excel 文件
    df = pd.read_excel(workbook_path)

    # 增加来源文件名字段，方便后续追踪数据来源
    df["source_file"] = workbook_path.name

    dataframes.append(df)

# 合并多个 DataFrame 为一个总表
merged_df = pd.concat(dataframes, ignore_index=True)

# 打印前 5 行
print(merged_df.head())
```

运行后会看到合并总表的前 5 行，并包含 `source_file` 来源列。

## 如何清洗 Excel 数据并按规则生成汇总结果？

清洗 Excel 数据并生成汇总结果，关键是先把字段、类型和重复记录处理一致，再用 `groupby`（按某列分组统计）计算指标。

### 清洗前后对比

| 项目 | 清洗前 | 清洗后 |
|---|---|---|
| 空值 | 部门为空、金额为空 | 填为“未知”或 0 |
| 字段名 | “销售部门”“部门”混用 | 统一为“部门” |
| 金额 | `"1,200"` 文本 | 转为数字 1200 |
| 重复记录 | 同一订单重复出现 | 按订单号删除 |

常见汇总规则包括：按部门、人员、产品、日期分组；统计订单数、销售额合计、平均值、最大值。

这段代码会清洗合并后的数据，并按部门生成销售额合计和订单数量：

```python
import pandas as pd

# 最小可复现示例：模拟多个 Excel 合并后的数据
df = pd.DataFrame({
    "销售部门": ["华东", "华东", None, "华南", "华南"],
    "订单号": ["A001", "A001", "A002", "A003", "A004"],
    "销售额": ["1,200", "1,200", "800", None, "2,500"],
    "下单日期": ["2024/01/01", "2024-01-01", "2024/01/02", "错误日期", "2024/01/03"]
})

# 统一字段名称，避免不同文件表头不一致
df = df.rename(columns={"销售部门": "部门"})

# 处理空值
df["部门"] = df["部门"].fillna("未知")
df["销售额"] = df["销售额"].fillna("0")

# 金额字段去掉逗号并转为数字
df["销售额"] = pd.to_numeric(df["销售额"].str.replace(",", ""), errors="coerce").fillna(0)

# 日期转为标准日期格式，错误日期变为空值
df["下单日期"] = pd.to_datetime(df["下单日期"], errors="coerce")

# 删除重复订单
df = df.drop_duplicates(subset=["订单号"])

# 按部门汇总销售额合计和订单数量
summary = df.groupby("部门", as_index=False).agg(
    销售额合计=("销售额", "sum"),
    订单数量=("订单号", "count"),
    平均销售额=("销售额", "mean"),
    最大销售额=("销售额", "max")
)

print(summary)
```

运行结果是一个 `summary` DataFrame，可直接用于导出 Excel 汇总报表。

## 如何把汇总结果导出为 Excel 报表并设置格式？

把汇总结果导出为 Excel 报表，推荐用 `pandas.to_excel` 写入数据，再用 `openpyxl` 设置列宽、冻结首行和表头加粗。`DataFrame`（二维表格数据结构）可通过 `ExcelWriter` 写入同一个 `report.xlsx` 的多个工作表。

| 输出内容 | 工作表名 | 用途 |
|---|---|---|
| 合并明细表 | 明细表 | 保存原始合并数据 |
| 部门汇总表 | 汇总表 | 按部门统计金额 |
| 异常数据表 | 异常数据 | 记录金额为空或小于 0 的行 |
| 生成时间 | 汇总表 A1 | 标记报表生成日期 |

这段代码会生成明细表和汇总表，并设置基础 Excel 样式：

```python
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# 1. 构造最小可运行示例数据
detail_df = pd.DataFrame({
    "部门": ["销售部", "销售部", "技术部", "技术部", "财务部"],
    "员工": ["张三", "李四", "王五", "赵六", "钱七"],
    "金额": [1200, 1800, 2600, 3100, 900],
    "日期": ["2026-06-01", "2026-06-01", "2026-06-02", "2026-06-02", "2026-06-03"]
})

# 2. 按部门汇总金额
summary_df = (
    detail_df.groupby("部门", as_index=False)["金额"]
    .sum()
    .rename(columns={"金额": "合计金额"})
)

report_path = Path("report.xlsx")

# 3. 把多个 DataFrame 写入同一个 Excel 文件的不同工作表
with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
    detail_df.to_excel(writer, sheet_name="明细表", index=False)
    summary_df.to_excel(writer, sheet_name="汇总表", index=False, startrow=2)

# 4. 使用 openpyxl 设置格式
workbook = load_workbook(report_path)

for worksheet in workbook.worksheets:
    # 冻结首行，滚动时表头保持可见
    worksheet.freeze_panes = "A2"

    # 设置表头加粗、背景色、居中
    for cell in worksheet[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
        cell.alignment = Alignment(horizontal="center")

    # 根据内容自动估算列宽
    for column_cells in worksheet.columns:
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))
        worksheet.column_dimensions[column_letter].width = max_length + 4

# 5. 在汇总表写入生成时间
summary_sheet = workbook["汇总表"]
summary_sheet["A1"] = f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
summary_sheet["A1"].font = Font(bold=True)

workbook.save(report_path)
print(f"报表已生成：{report_path.resolve()}")
```

运行后，当前目录会出现 `report.xlsx`，其中包含“明细表”和“汇总表”，并已完成列宽、冻结首行和表头样式设置。

## 要点回顾

- Python 批量处理 Excel 的标准流程是读取文件、合并数据、清洗字段、分组汇总、导出报表。
- 新手处理 Excel 自动化时，优先使用 pandas 完成数据分析，再用 openpyxl 做格式美化。
- 只要原始 Excel 表头相对统一，就可以用同一套脚本反复生成日报、周报或月报。

## 常见问答(FAQ)

**Q:Python 处理 Excel 应该用 pandas 还是 openpyxl？**

A:优先用 pandas 处理数据，用 openpyxl 调整格式。pandas 适合读取、筛选、合并、分组统计，例如汇总每个部门销售额；openpyxl 更适合设置列宽、字体、冻结首行等 Excel 样式。

**Q:多个 Excel 表头不完全一样怎么办？**

A:先统一字段名，再合并数据。可以建立一张字段映射表，例如把“客户名称”“客户名”“公司名称”都改成“客户”。缺失列可补空值，多余列可删除，避免 concat 后出现大量重复字段。

**Q:可以自动生成带多个工作表的报表吗？**

A:可以，用 pandas 的 ExcelWriter 一次写入多个工作表。例如把“明细数据”“部门汇总”“异常记录”分别导出到同一个 xlsx 文件中，方便查看和分发，再用 openpyxl 做简单排版。

**Q:Python 能处理很大的 Excel 文件吗？**

A:能处理，但要看文件大小和内存。几十万行数据通常可用 pandas 读取；如果超过百万行，建议改用 CSV、分块读取，或先写入数据库再统计。Excel 本身单表也有约 104 万行限制。

**Q:脚本能每天自动运行生成报表吗？**

A:能，可以把脚本放到定时任务里运行。Windows 可用任务计划程序，Linux 或服务器可用 crontab，例如每天早上 8 点读取前一天文件，生成汇总报表并保存到指定文件夹。

<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "Article",
  "headline": "如何用 Python 自动批量处理 Excel 表格并生成汇总报表？新手完整流程",
  "description": "用 Python 批量处理 Excel 的核心流程是：读取多个文件、清洗和合并数据、按规则统计汇总，最后导出新的报表。新手建议优先使用 pandas 处理数据，用 pathlib 扫描文件，用 openpyxl 辅助设置 Excel 格式，就能把重复的人工统计工作自动化。",
  "keywords": [
    "Python 自动化",
    "Python 处理 Excel",
    "批量生成报表"
  ],
  "datePublished": "2026-06-04T09:34:34",
  "dateModified": "2026-06-04",
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
      "name": "Python 处理 Excel 应该用 pandas 还是 openpyxl？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "优先用 pandas 处理数据，用 openpyxl 调整格式。pandas 适合读取、筛选、合并、分组统计，例如汇总每个部门销售额；openpyxl 更适合设置列宽、字体、冻结首行等 Excel 样式。"
      }
    },
    {
      "@type": "Question",
      "name": "多个 Excel 表头不完全一样怎么办？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "先统一字段名，再合并数据。可以建立一张字段映射表，例如把“客户名称”“客户名”“公司名称”都改成“客户”。缺失列可补空值，多余列可删除，避免 concat 后出现大量重复字段。"
      }
    },
    {
      "@type": "Question",
      "name": "可以自动生成带多个工作表的报表吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "可以，用 pandas 的 ExcelWriter 一次写入多个工作表。例如把“明细数据”“部门汇总”“异常记录”分别导出到同一个 xlsx 文件中，方便查看和分发，再用 openpyxl 做简单排版。"
      }
    },
    {
      "@type": "Question",
      "name": "Python 能处理很大的 Excel 文件吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "能处理，但要看文件大小和内存。几十万行数据通常可用 pandas 读取；如果超过百万行，建议改用 CSV、分块读取，或先写入数据库再统计。Excel 本身单表也有约 104 万行限制。"
      }
    },
    {
      "@type": "Question",
      "name": "脚本能每天自动运行生成报表吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "能，可以把脚本放到定时任务里运行。Windows 可用任务计划程序，Linux 或服务器可用 crontab，例如每天早上 8 点读取前一天文件，生成汇总报表并保存到指定文件夹。"
      }
    }
  ]
}
</script>
<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "HowTo",
  "name": "如何用 Python 自动批量处理 Excel 表格并生成汇总报表？新手完整流程",
  "description": "用 Python 批量处理 Excel 的核心流程是：读取多个文件、清洗和合并数据、按规则统计汇总，最后导出新的报表。新手建议优先使用 pandas 处理数据，用 pathlib 扫描文件，用 openpyxl 辅助设置 Excel 格式，就能把重复的人工统计工作自动化。",
  "step": [
    {
      "@type": "HowToStep",
      "name": "为什么适合用 Python 自动化处理 Excel 报表？",
      "text": "Python 适合自动化处理 Excel 报表，因为它能把“逐个打开文件、复制粘贴、手工统计”变成可重复运行的脚本。人工处理销售日报、库存表、考勤表、财务明细、项目进度表时，常见问题是耗时、漏算、列名不一致、日期格式混乱。"
    },
    {
      "@type": "HowToStep",
      "name": "如何准备 Python 处理 Excel 的运行环境？",
      "text": "准备环境只需安装 Python 3.10+、pandas 和 openpyxl，并建立固定目录。pandas（数据处理库）负责合并统计，openpyxl（xlsx 读写库）负责 Excel 文件读写和格式。"
    },
    {
      "@type": "HowToStep",
      "name": "如何用 Python 批量读取多个 Excel 文件？",
      "text": "用 Python 批量读取多个 Excel 文件，可以用 `pathlib` 扫描 `input` 文件夹中的所有 `.xlsx`，再用 `pandas.read_excel` 逐个读成 DataFrame（pandas 的二维表格数据结构）并合并。"
    }
  ],
  "tool": [
    {
      "@type": "HowToTool",
      "name": "Python 3.12、pandas 2.x、openpyxl 3.x、Windows/macOS/Linux 均可"
    }
  ]
}
</script>
