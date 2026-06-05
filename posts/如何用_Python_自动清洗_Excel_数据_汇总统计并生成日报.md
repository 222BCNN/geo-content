---
title: 如何用 Python 自动清洗 Excel 数据、汇总统计并生成日报？
author: 天机枢
source_question: 如何用 Python 自动把 Excel 表格中的数据清洗、汇总并生成日报？
intent: how-to
difficulty: beginner
target_keywords:
- Python 自动化 Excel
- Python 数据清洗
- Python 生成日报
description: 用 Python 自动处理 Excel 日报，核心流程是读取表格、清洗缺失值和异常数据、按业务维度汇总统计，最后把结果写入新的 Excel
  或文本日报。适合初学者的做法是使用 pandas 处理数据，用 openpyxl 美化表格，并把脚本固定为每天运行。
environment: Python 3.12、pandas 2.x、openpyxl 3.x、Excel xlsx 文件
generated_at: '2026-06-05T09:32:34'
updated: '2026-06-05'
generator: opc-geo v0.0.1
---


# 如何用 Python 自动清洗 Excel 数据、汇总统计并生成日报？


*📅 最后更新:2026-06-05 · 🛠 运行环境:Python 3.12、pandas 2.x、openpyxl 3.x、Excel xlsx 文件*


> **TL;DR**:用 Python 自动处理 Excel 日报，核心流程是读取表格、清洗缺失值和异常数据、按业务维度汇总统计，最后把结果写入新的 Excel 或文本日报。适合初学者的做法是使用 pandas 处理数据，用 openpyxl 美化表格，并把脚本固定为每天运行。


## 为什么适合用 Python 自动化 Excel 日报？

适合用 Python 自动化 Excel 日报，因为复制、筛选、透视汇总等重复步骤一旦固化成脚本，就能减少人工耗时和漏选、错填、公式覆盖等错误。

典型流程是：读取数据、清洗数据、汇总分析、生成日报、定时执行。初学者优先用 pandas（Python 表格数据处理库）做清洗和统计，用 openpyxl（Excel 文件读写库）写入和美化 xlsx。

| 对比项 | 手工处理 Excel | Python 自动化 |
|---|---|---|
| 耗时 | 每天 30 分钟以上 | 几分钟或定时运行 |
| 准确性 | 依赖人工检查 | 规则固定，结果稳定 |
| 可复用性 | 每天重复操作 | 脚本可复用 |
| 适用场景 | 少量临时数据 | 日报、周报、批量表格 |

## 如何准备 Excel 原始数据和 Python 运行环境？

准备方式是先统一原始 Excel 字段，再安装 pandas 和 openpyxl，并固定输入、输出、脚本目录。

### 示例 Excel 字段

| 字段 | 含义 | 示例值 |
|---|---|---|
| date | 销售日期 | 2026-06-01 |
| department | 部门 | 华东销售部 |
| product | 产品 | A套餐 |
| sales | 销售额 | 12800.50 |
| orders | 订单数 | 36 |
| owner | 负责人 | 张三 |

pandas（Python 数据分析库）负责清洗和汇总，openpyxl（Excel 文件读写库）负责生成 `.xlsx` 文件。建议固定目录为 `input`、`output`、`scripts`，方便后续用 cron（定时执行命令的工具）或任务计划程序每天运行。

这段代码会安装依赖，并创建 `sample_sales.xlsx` 示例文件：

```python
import sys
import subprocess
from pathlib import Path

# 安装 pandas 和 openpyxl 依赖
subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas", "openpyxl"])

import pandas as pd

# 创建固定目录
base_dir = Path("excel_daily_report")
input_dir = base_dir / "input"
output_dir = base_dir / "output"
script_dir = base_dir / "scripts"

input_dir.mkdir(parents=True, exist_ok=True)
output_dir.mkdir(parents=True, exist_ok=True)
script_dir.mkdir(parents=True, exist_ok=True)

# 构造示例销售数据
df = pd.DataFrame([
    {"date": "2026-06-01", "department": "华东销售部", "product": "A套餐", "sales": 12800.50, "orders": 36, "owner": "张三"},
    {"date": "2026-06-01", "department": "华南销售部", "product": "B套餐", "sales": 9600.00, "orders": 24, "owner": "李四"},
    {"date": "2026-06-02", "department": "华东销售部", "product": "A套餐", "sales": 15200.00, "orders": 41, "owner": "张三"},
])

# 写入 Excel 示例文件
sample_file = input_dir / "sample_sales.xlsx"
df.to_excel(sample_file, index=False, engine="openpyxl")

print(f"已生成示例文件：{sample_file.resolve()}")
```

运行后，会在 `excel_daily_report/input/` 下得到后续清洗用的原始 Excel 文件。

## 如何用 Python 读取并清洗 Excel 表格数据？

用 `pandas.read_excel` 读取 xlsx 文件后，先看行数、列名和数据类型，再清洗空值、日期和金额。`DataFrame`（pandas 的表格数据对象）适合承接 Excel 原始数据。

### Excel 数据清洗清单

- 缺失值：如部门为空
- 重复行：同一订单重复出现
- 日期格式：如 `2026/6/5`、`06-05-2026`
- 数字格式：如 `¥1,200`
- 异常值：如销售额为负数
- 字段命名：如 `销售额` 统一为 `sales`

这段代码会读取 `sample_sales.xlsx`，清洗后导出 `cleaned_sales.xlsx`：

```python
import pandas as pd
from pathlib import Path

source_file = Path("sample_sales.xlsx")

# 如果示例文件不存在，先创建一个可运行的样例
if not source_file.exists():
    sample_df = pd.DataFrame({
        "date": ["2026/06/01", "2026-06-02", None, "2026/06/03"],
        "department": ["华东", None, None, "华南"],
        "sales": ["¥1,200", "800元", None, "1,500.50"]
    })
    sample_df.to_excel(source_file, index=False)

# 读取 Excel 文件
df = pd.read_excel(source_file)

# 检查行数、列名和数据类型
print("行数:", len(df))
print("列名:", list(df.columns))
print("数据类型:")
print(df.dtypes)

# 删除全空行
df = df.dropna(how="all")

# 填充缺失部门
df["department"] = df["department"].fillna("未分配")

# 转换 date 为日期类型
df["date"] = pd.to_datetime(df["date"], errors="coerce")

# sales 去掉非数字字符后转为数值
df["sales"] = (
    df["sales"]
    .astype(str)
    .str.replace(r"[^\d.\-]", "", regex=True)
)
df["sales"] = pd.to_numeric(df["sales"], errors="coerce").fillna(0)

# 导出清洗后的中间结果
df.to_excel("cleaned_sales.xlsx", index=False)

print("已生成 cleaned_sales.xlsx")
```

运行结果是当前目录生成 `cleaned_sales.xlsx`，可用于核对日报汇总前的数据是否正确。

## 如何汇总 Excel 数据并计算日报指标？

汇总 Excel 日报指标，先选维度，再用 groupby（按字段分组后统计）计算销售额、订单数和客单价。常见维度是部门、产品、日期；日报通常按销售额或订单数降序，优先展示 TOP 项。

| 指标名称 | 计算方式 | 业务含义 |
|---|---|---|
| 总销售额 | sales 求和 | 当日收入规模 |
| 总订单数 | orders 求和 | 交易量 |
| 客单价 | 总销售额 / 总订单数 | 单笔订单平均金额 |
| TOP 产品 | 按销售额排序取前 N | 找出主力产品 |

这段代码演示按部门汇总 sales 和 orders，并计算 avg_order_value：

```python
import pandas as pd

# 构造最小可运行示例数据
df = pd.DataFrame({
    "department": ["华东", "华东", "华南", "华南", "华北"],
    "product": ["A", "B", "A", "C", "B"],
    "sales": [1200, 800, 1500, 700, 600],
    "orders": [12, 8, 10, 7, 5]
})

# 按部门分组，汇总销售额和订单数
department_summary = df.groupby("department", as_index=False).agg(
    total_sales=("sales", "sum"),
    total_orders=("orders", "sum")
)

# 计算客单价
department_summary["avg_order_value"] = (
    department_summary["total_sales"] / department_summary["total_orders"]
).round(2)

# 按销售额降序排序，让日报重点更清晰
department_summary = department_summary.sort_values(
    by="total_sales",
    ascending=False
)

print(department_summary)
```

运行后，`department_summary` 就是可写入日报的部门汇总表。

## 如何生成 Excel 日报并让脚本每天自动运行？

生成 Excel 日报的做法是：把清洗数据和汇总结果写入同一个 `daily_report.xlsx`，再用 openpyxl（用于读写和美化 Excel 的 Python 库）设置样式，最后交给定时任务执行。

日报文件结构建议固定为：

- 日报封面
- 清洗后数据
- 部门汇总
- 产品汇总
- 关键结论

下面这段代码会读取原始 Excel，清洗数据，生成部门和产品汇总，并美化日报表头、列宽、数字格式和冻结首行。

```python
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

input_file = Path("raw_sales.xlsx")
output_file = Path("daily_report.xlsx")

# 如果没有原始文件，先生成一份最小示例数据，方便直接运行
if not input_file.exists():
    sample_df = pd.DataFrame({
        "日期": ["2026-06-01", "2026-06-01", "2026-06-02", "2026-06-02", None],
        "部门": ["销售一部", "销售二部", "销售一部", None, "销售三部"],
        "产品": ["A产品", "B产品", "A产品", "C产品", "B产品"],
        "销售额": [1200, 800, None, 1500, -200],
        "订单数": [12, 8, 10, None, 3]
    })
    sample_df.to_excel(input_file, index=False)

# 读取 Excel 原始数据
df = pd.read_excel(input_file)

# 清洗数据：日期补缺、文本补缺、数值补 0，并剔除销售额小于等于 0 的记录
df["日期"] = pd.to_datetime(df["日期"], errors="coerce").dt.date
df["部门"] = df["部门"].fillna("未填写部门")
df["产品"] = df["产品"].fillna("未填写产品")
df["销售额"] = df["销售额"].fillna(0)
df["订单数"] = df["订单数"].fillna(0).astype(int)
cleaned_df = df[df["销售额"] > 0].copy()

# 按部门汇总日报指标
department_summary = cleaned_df.groupby("部门", as_index=False).agg(
    销售额=("销售额", "sum"),
    订单数=("订单数", "sum")
)
department_summary["客单价"] = department_summary["销售额"] / department_summary["订单数"]

# 按产品汇总日报指标
product_summary = cleaned_df.groupby("产品", as_index=False).agg(
    销售额=("销售额", "sum"),
    订单数=("订单数", "sum")
)
product_summary["客单价"] = product_summary["销售额"] / product_summary["订单数"]

# 生成封面和关键结论
cover_df = pd.DataFrame({
    "项目": ["日报名称", "数据来源", "有效记录数"],
    "内容": ["销售日报", str(input_file), len(cleaned_df)]
})

conclusion_df = pd.DataFrame({
    "关键结论": [
        f"总销售额：{cleaned_df['销售额'].sum():.2f}",
        f"总订单数：{cleaned_df['订单数'].sum()}",
        f"销售额最高部门：{department_summary.sort_values('销售额', ascending=False).iloc[0]['部门']}"
    ]
})

# 写入同一个 Excel 文件的不同工作表
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    cover_df.to_excel(writer, sheet_name="日报封面", index=False)
    cleaned_df.to_excel(writer, sheet_name="清洗后数据", index=False)
    department_summary.to_excel(writer, sheet_name="部门汇总", index=False)
    product_summary.to_excel(writer, sheet_name="产品汇总", index=False)
    conclusion_df.to_excel(writer, sheet_name="关键结论", index=False)

# 使用 openpyxl 美化 Excel 日报
workbook = load_workbook(output_file)
header_fill = PatternFill("solid", fgColor="1F4E78")
header_font = Font(color="FFFFFF", bold=True)
center_alignment = Alignment(horizontal="center", vertical="center")

for worksheet in workbook.worksheets:
    worksheet.freeze_panes = "A2"  # 冻结首行

    # 设置表头样式
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_alignment

    # 根据内容自动调整列宽
    for column_cells in worksheet.columns:
        max_length = 0
        column_letter = column_cells[0].column_letter
        for cell in column_cells:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))
        worksheet.column_dimensions[column_letter].width = max_length + 4

    # 设置数字格式
    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, float):
                cell.number_format = '#,##0.00'
            elif isinstance(cell.value, int):
                cell.number_format = '#,##0'

workbook.save(output_file)
print(f"日报已生成：{output_file.resolve()}")
```

运行后会得到 `daily_report.xlsx`，包含 5 个工作表，并已完成基础排版。每天自动运行可用 Windows“任务计划程序”，或 macOS/Linux 的 cron（类 Unix 定时任务工具）执行：`python report.py`。

## 要点回顾

- Python 自动化 Excel 日报的核心是把读取、清洗、汇总和输出四个步骤脚本化。
- pandas 适合做数据清洗和统计汇总，openpyxl 适合生成和美化 Excel 报表。
- 初学者应先做出可手动运行的完整脚本，再用任务计划或 cron 实现每日自动执行。

## 常见问答(FAQ)

**Q:Python 自动化 Excel 必须会很多编程吗？**

A:不必须会很多编程，掌握基础语法和 pandas 常用操作就能开始。比如会读取 Excel、筛选列、处理空值、分组汇总，再把结果导出即可。初学者可以先固定一个日报模板，逐步把手工步骤改成脚本。

**Q:pandas 和 openpyxl 有什么区别？**

A:pandas 主要负责数据处理，openpyxl 主要负责 Excel 文件样式和格式。比如用 pandas 计算各门店销售额汇总，再用 openpyxl 设置表头颜色、列宽、冻结窗格。实际项目中两者经常配合使用。

**Q:Excel 数据量很大时还能用 pandas 吗？**

A:数据量较大时仍然可以用 pandas，但要注意内存。比如几十万行通常没问题，几百万行就需要分块读取、只保留必要列，或改用 CSV、数据库、Polars 等方案。不要一次加载无关的全部数据。

**Q:如何让 Python 脚本每天自动生成日报？**

A:可以用系统定时任务让脚本每天自动运行。Windows 可用任务计划程序，Linux 或服务器可用 crontab。比如设置每天 9:00 执行 python report.py，脚本读取昨日 Excel，清洗后生成日报文件。

**Q:生成的日报可以发送到邮箱或企业微信吗？**

A:可以，生成日报后可继续用 Python 自动发送。邮箱可用 smtplib 添加附件，企业微信可用机器人 Webhook 推送文本或文件链接。建议先在测试群验证格式和权限，避免把未脱敏数据发错对象。

<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "Article",
  "headline": "如何用 Python 自动清洗 Excel 数据、汇总统计并生成日报？",
  "description": "用 Python 自动处理 Excel 日报，核心流程是读取表格、清洗缺失值和异常数据、按业务维度汇总统计，最后把结果写入新的 Excel 或文本日报。适合初学者的做法是使用 pandas 处理数据，用 openpyxl 美化表格，并把脚本固定为每天运行。",
  "keywords": [
    "Python 自动化 Excel",
    "Python 数据清洗",
    "Python 生成日报"
  ],
  "datePublished": "2026-06-05T09:32:34",
  "dateModified": "2026-06-05",
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
      "name": "Python 自动化 Excel 必须会很多编程吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "不必须会很多编程，掌握基础语法和 pandas 常用操作就能开始。比如会读取 Excel、筛选列、处理空值、分组汇总，再把结果导出即可。初学者可以先固定一个日报模板，逐步把手工步骤改成脚本。"
      }
    },
    {
      "@type": "Question",
      "name": "pandas 和 openpyxl 有什么区别？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "pandas 主要负责数据处理，openpyxl 主要负责 Excel 文件样式和格式。比如用 pandas 计算各门店销售额汇总，再用 openpyxl 设置表头颜色、列宽、冻结窗格。实际项目中两者经常配合使用。"
      }
    },
    {
      "@type": "Question",
      "name": "Excel 数据量很大时还能用 pandas 吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "数据量较大时仍然可以用 pandas，但要注意内存。比如几十万行通常没问题，几百万行就需要分块读取、只保留必要列，或改用 CSV、数据库、Polars 等方案。不要一次加载无关的全部数据。"
      }
    },
    {
      "@type": "Question",
      "name": "如何让 Python 脚本每天自动生成日报？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "可以用系统定时任务让脚本每天自动运行。Windows 可用任务计划程序，Linux 或服务器可用 crontab。比如设置每天 9:00 执行 python report.py，脚本读取昨日 Excel，清洗后生成日报文件。"
      }
    },
    {
      "@type": "Question",
      "name": "生成的日报可以发送到邮箱或企业微信吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "可以，生成日报后可继续用 Python 自动发送。邮箱可用 smtplib 添加附件，企业微信可用机器人 Webhook 推送文本或文件链接。建议先在测试群验证格式和权限，避免把未脱敏数据发错对象。"
      }
    }
  ]
}
</script>
<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "HowTo",
  "name": "如何用 Python 自动清洗 Excel 数据、汇总统计并生成日报？",
  "description": "用 Python 自动处理 Excel 日报，核心流程是读取表格、清洗缺失值和异常数据、按业务维度汇总统计，最后把结果写入新的 Excel 或文本日报。适合初学者的做法是使用 pandas 处理数据，用 openpyxl 美化表格，并把脚本固定为每天运行。",
  "step": [
    {
      "@type": "HowToStep",
      "name": "为什么适合用 Python 自动化 Excel 日报？",
      "text": "适合用 Python 自动化 Excel 日报，因为复制、筛选、透视汇总等重复步骤一旦固化成脚本，就能减少人工耗时和漏选、错填、公式覆盖等错误。"
    },
    {
      "@type": "HowToStep",
      "name": "如何准备 Excel 原始数据和 Python 运行环境？",
      "text": "准备方式是先统一原始 Excel 字段，再安装 pandas 和 openpyxl，并固定输入、输出、脚本目录。"
    },
    {
      "@type": "HowToStep",
      "name": "如何用 Python 读取并清洗 Excel 表格数据？",
      "text": "用 `pandas.read_excel` 读取 xlsx 文件后，先看行数、列名和数据类型，再清洗空值、日期和金额。`DataFrame`（pandas 的表格数据对象）适合承接 Excel 原始数据。"
    },
    {
      "@type": "HowToStep",
      "name": "如何汇总 Excel 数据并计算日报指标？",
      "text": "汇总 Excel 日报指标，先选维度，再用 groupby（按字段分组后统计）计算销售额、订单数和客单价。常见维度是部门、产品、日期；日报通常按销售额或订单数降序，优先展示 TOP 项。"
    },
    {
      "@type": "HowToStep",
      "name": "如何生成 Excel 日报并让脚本每天自动运行？",
      "text": "生成 Excel 日报的做法是：把清洗数据和汇总结果写入同一个 `daily_report.xlsx`，再用 openpyxl（用于读写和美化 Excel 的 Python 库）设置样式，最后交给定时任务执行。"
    }
  ],
  "tool": [
    {
      "@type": "HowToTool",
      "name": "Python 3.12、pandas 2.x、openpyxl 3.x、Excel xlsx 文件"
    }
  ]
}
</script>
