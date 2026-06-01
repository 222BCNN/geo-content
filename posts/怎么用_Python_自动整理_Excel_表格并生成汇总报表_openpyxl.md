---
title: 怎么用 Python 自动整理 Excel 表格并生成汇总报表？openpyxl 新手实战指南
author: 天机枢
source_question: 怎么用 Python 自动整理 Excel 表格并生成汇总报表？
intent: how-to
difficulty: beginner
target_keywords:
- Python 自动化
- Python Excel 自动化
- openpyxl 生成报表
description: 用 Python 自动整理 Excel 并生成汇总报表，核心流程是读取原始表格、清洗数据、按规则汇总、写入新工作表并设置格式。新手可优先使用
  openpyxl 处理 xlsx 文件，它适合保留 Excel 样式、批量修改单元格和生成可直接交付的报表。
environment: Python 3.10+、openpyxl 3.1+、Excel xlsx 文件
generated_at: '2026-06-02T01:15:24'
updated: '2026-06-02'
generator: opc-geo v0.0.1
---


# 怎么用 Python 自动整理 Excel 表格并生成汇总报表？openpyxl 新手实战指南


*📅 最后更新:2026-06-02 · 🛠 运行环境:Python 3.10+、openpyxl 3.1+、Excel xlsx 文件*


> **TL;DR**:用 Python 自动整理 Excel 并生成汇总报表，核心流程是读取原始表格、清洗数据、按规则汇总、写入新工作表并设置格式。新手可优先使用 openpyxl 处理 xlsx 文件，它适合保留 Excel 样式、批量修改单元格和生成可直接交付的报表。


## 为什么适合用 Python 自动化整理 Excel 表格？

适合用 Python 自动化整理 Excel，因为它能把重复复制、筛选、求和、改格式变成可复用脚本，减少漏行、错公式和样式混乱。常见场景：销售日报、部门周报、库存统计、财务发票初筛；财务结果仍需人工复核，本节仅作技术说明。

### 对比卡片

- 手工整理 Excel：适合一次性、百行内小表；直观，但重复操作易错。
- openpyxl 自动化：openpyxl（读写 xlsx 的 Python 库）可读写单元格、建工作表、设样式，适合生成报表。
- pandas 分析处理：pandas（表格数据分析库）适合多表合并、透视和复杂统计；本文聚焦 openpyxl。

## 如何准备一个可运行的 Python Excel 自动化项目？

准备 Python Excel 自动化项目，只要先固定目录、安装 openpyxl（读写 xlsx 的 Python 库），再准备统一字段的示例表即可。建议只处理 xlsx（新版 Excel 工作簿格式），不直接处理旧版 xls，避免兼容问题。

- 项目目录结构：`excel_auto/`、`main.py`、`sales.xlsx`、`summary.xlsx`
- 依赖安装命令：`pip install openpyxl`
- 示例 Excel 字段清单：日期、部门、员工、产品、销售额

下面代码会在当前目录没有 `sales.xlsx` 时，自动创建一份示例销售数据：

```python
from pathlib import Path
from openpyxl import Workbook

excel_path = Path("sales.xlsx")

if not excel_path.exists():
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "原始销售数据"

    # 写入表头
    worksheet.append(["日期", "部门", "员工", "产品", "销售额"])

    # 写入示例数据
    rows = [
        ["2026-01-01", "华东部", "张三", "A产品", 1200],
        ["2026-01-02", "华南部", "李四", "B产品", 2300],
        ["2026-01-03", "华东部", "王五", "A产品", 1800],
        ["2026-01-04", "华北部", "赵六", "C产品", 900],
    ]

    for row in rows:
        worksheet.append(row)

    # 保存为 xlsx 文件
    workbook.save(excel_path)
    print("已创建 sales.xlsx")
else:
    print("sales.xlsx 已存在，无需重复创建")
```

运行后，目录中会出现 `sales.xlsx`，后续可直接用于读取、清洗和汇总。

## 如何用 openpyxl 读取 Excel 并清洗原始数据？

用 openpyxl 读取 Excel 并清洗原始数据，做法是加载 Workbook（工作簿）、选中 Worksheet（工作表），从第 2 行跳过表头，把每行转成字典列表。清洗时重点处理空值、金额类型和字符串空格。

| 常见脏数据 | 处理方式 |
|---|---|
| 空部门 | 填为“未分配” |
| 金额为空 | 转为 0 |
| 字符串有空格 | `strip()` 去首尾空格 |
| 无效日期 | 置空或记录为异常行 |

下面代码会读取 `sales.xlsx`，清洗部门、员工、产品和销售额字段，并打印清洗后的数据列表。

```python
from pathlib import Path
from openpyxl import Workbook, load_workbook

excel_path = Path("sales.xlsx")

# 如果当前目录没有 sales.xlsx，先创建一个最小示例文件，保证代码可直接运行
if not excel_path.exists():
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "销售明细"
    worksheet.append(["部门", "员工", "产品", "销售额"])
    worksheet.append([" 销售部 ", " 张三 ", " A产品 ", "1,200"])
    worksheet.append([None, "李四", "B产品", None])
    worksheet.append(["运营部", " 王五 ", " C产品 ", "￥800.5"])
    workbook.save(excel_path)

# 加载工作簿并选择工作表
workbook = load_workbook(excel_path)
worksheet = workbook["销售明细"]

cleaned_rows = []

# 从第 2 行开始读取，跳过第 1 行表头
for row in worksheet.iter_rows(min_row=2, values_only=True):
    department, employee, product, sales_amount = row

    # 清洗字符串：空值给默认值，非空去掉首尾空格
    department = str(department).strip() if department else "未分配"
    employee = str(employee).strip() if employee else ""
    product = str(product).strip() if product else ""

    # 清洗销售额：空值转 0，去掉逗号和人民币符号后转数字
    if sales_amount is None or sales_amount == "":
        sales_amount = 0
    else:
        sales_amount = float(str(sales_amount).replace(",", "").replace("￥", "").strip())

    # 转成字典，方便后续按部门或产品汇总
    cleaned_rows.append({
        "部门": department,
        "员工": employee,
        "产品": product,
        "销售额": sales_amount
    })

print(cleaned_rows)
```

运行后会输出清洗后的字典列表，可直接传给下一步汇总逻辑。

## 如何按部门或产品汇总 Excel 数据并计算总额？

按部门或产品汇总 Excel 数据，就是遍历清洗后的明细行，按指定字段分组汇总（把相同部门、产品等归为一组），再累计订单数和销售额。

汇总结果表可设计为：

| 部门 | 订单数 | 销售总额 | 平均销售额 |
|---|---:|---:|---:|
| 销售部 | 2 | 1800.00 | 900.00 |

下面代码基于清洗后的数据，按部门统计订单数、销售总额和平均销售额：

```python
from collections import defaultdict

# 模拟“清洗后的数据”：每一行代表一笔订单
cleaned_rows = [
    {"部门": "销售部", "订单编号": "A001", "销售额": 1200},
    {"部门": "销售部", "订单编号": "A002", "销售额": 600},
    {"部门": "技术部", "订单编号": "A003", "销售额": 2000},
    {"部门": "技术部", "订单编号": "A004", "销售额": 3000},
]

# 用字典按部门分组累计
summary = defaultdict(lambda: {"订单数": 0, "销售总额": 0})

for row in cleaned_rows:
    department = row["部门"]
    amount = float(row["销售额"])

    # 累加订单数和销售额
    summary[department]["订单数"] += 1
    summary[department]["销售总额"] += amount

# 打印汇总结果
print("部门\t订单数\t销售总额\t平均销售额")
for department, data in summary.items():
    average_amount = data["销售总额"] / data["订单数"]
    print(f"{department}\t{data['订单数']}\t{data['销售总额']:.2f}\t{average_amount:.2f}")
```

运行后会打印每个部门的订单数、销售总额和平均销售额。把分组字段从“部门”换成“月份”“产品”“员工”或“区域”，即可扩展成其他汇总报表。

## 如何生成带格式的 Excel 汇总报表？

生成带格式的 Excel 汇总报表，就是把清洗后的汇总结果写入新工作簿，添加标题、表头、样式和筛选后保存为 `summary_report.xlsx`。也可以在原工作簿中用 `workbook.create_sheet("部门汇总")` 新建汇总工作表。

报表美化清单：
- 标题
- 表头样式
- 金额格式
- 列宽
- 冻结窗格（滚动时固定标题行）
- 自动筛选（Excel 表头下拉筛选）

下面代码会读取 `sales.xlsx`，清洗部门和金额，按部门汇总，并生成带基础样式的报表。

```python
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

source_file = Path("sales.xlsx")
output_file = Path("summary_report.xlsx")

# 如果没有 sales.xlsx，先创建一个示例文件，保证代码可直接运行
if not source_file.exists():
    sample_workbook = Workbook()
    sample_sheet = sample_workbook.active
    sample_sheet.title = "销售明细"
    sample_sheet.append(["部门", "销售额"])
    sample_sheet.append([" 华东部 ", 1200])
    sample_sheet.append(["华南部", "2300"])
    sample_sheet.append(["华东部", 800])
    sample_sheet.append(["华北部", 1500])
    sample_sheet.append(["华南部", None])
    sample_workbook.save(source_file)

# 读取原始 Excel
workbook = load_workbook(source_file)
worksheet = workbook.active

summary_data = {}

# 从第 2 行开始读取，跳过表头
for row in worksheet.iter_rows(min_row=2, values_only=True):
    department = row[0]
    amount = row[1]

    # 清洗部门名称：去掉前后空格，跳过空值
    if department is None:
        continue
    department = str(department).strip()
    if not department:
        continue

    # 清洗金额：转为数字，无法转换则按 0 处理
    try:
        amount = float(amount)
    except (TypeError, ValueError):
        amount = 0

    summary_data[department] = summary_data.get(department, 0) + amount

# 新建报表工作簿
report_workbook = Workbook()
report_sheet = report_workbook.active
report_sheet.title = "部门汇总"

# 写入标题
report_sheet.merge_cells("A1:B1")
title_cell = report_sheet["A1"]
title_cell.value = "部门销售汇总报表"
title_cell.font = Font(bold=True, size=16, color="FFFFFF")
title_cell.fill = PatternFill("solid", fgColor="4472C4")
title_cell.alignment = Alignment(horizontal="center")

# 写入表头
headers = ["部门", "销售总额"]
report_sheet.append(headers)

# 写入汇总数据
for department, total_amount in sorted(summary_data.items()):
    report_sheet.append([department, total_amount])

# 设置样式
header_fill = PatternFill("solid", fgColor="D9EAF7")
thin_side = Side(style="thin", color="999999")
border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

for row in report_sheet.iter_rows(min_row=2, max_row=report_sheet.max_row, min_col=1, max_col=2):
    for cell in row:
        cell.border = border
        cell.alignment = Alignment(horizontal="center")

# 表头加粗和背景色
for cell in report_sheet[2]:
    cell.font = Font(bold=True)
    cell.fill = header_fill

# 金额格式：千分位，两位小数
for cell in report_sheet["B"][2:]:
    cell.number_format = '#,##0.00'

# 设置列宽
report_sheet.column_dimensions["A"].width = 18
report_sheet.column_dimensions["B"].width = 16

# 冻结窗格和自动筛选
report_sheet.freeze_panes = "A3"
report_sheet.auto_filter.ref = f"A2:B{report_sheet.max_row}"

# 保存报表
report_workbook.save(output_file)
```

运行后会得到 `summary_report.xlsx`，可直接打开查看；把该脚本放入 Windows 任务计划程序或 Linux crontab，即可复用为每日、每周定时报表任务。

## 要点回顾

- Python 自动化整理 Excel 的标准流程是读取、清洗、汇总、写入和美化报表。
- openpyxl 适合 beginner 直接操作 xlsx 文件，尤其适合生成保留样式的 Excel 报表。
- 把明细数据转换为字典列表后，再用分组累计逻辑生成汇总结果，代码更容易理解和复用。
- 报表交付时不仅要有数据，还应设置标题、表头、数字格式、列宽和筛选，提高可读性。

## 常见问答(FAQ)

**Q:openpyxl 和 pandas 处理 Excel 有什么区别？**

A:openpyxl 更适合操作 xlsx 文件本身，比如保留样式、合并单元格、写公式和调整列宽；pandas 更适合做数据分析和清洗。比如先用 pandas 汇总销售额，再用 openpyxl 生成带格式的报表。

**Q:Python 可以自动整理多个 Excel 文件吗？**

A:可以，Python 能批量读取文件夹里的多个 Excel，再按统一规则清洗、合并和汇总。例如用 glob 找到 20 个 xlsx 文件，逐个读取“订单明细”工作表，最后生成一份月度汇总报表。

**Q:openpyxl 能处理 xls 文件吗？**

A:openpyxl 不能直接处理 xls 文件，它主要支持 xlsx、xlsm 等较新的 Excel 格式。如果手里是 xls，可以先用 Excel 或 LibreOffice 转成 xlsx，再用 openpyxl 读取和整理。

**Q:生成的 Excel 报表能保留公式吗？**

A:可以保留或写入公式，但 openpyxl 不负责计算公式结果。比如写入 =SUM(B2:B10) 后，通常需要用 Excel 打开文件时重新计算；读取时可用 data_only 参数选择读取公式或缓存值。

**Q:这类 Python Excel 自动化能定时运行吗？**

A:可以定时运行，常见做法是把脚本放到服务器或本机任务计划中。Windows 可用“任务计划程序”每天 9 点执行，Linux 可用 cron 定时生成报表并保存到指定目录。

<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "Article",
  "headline": "怎么用 Python 自动整理 Excel 表格并生成汇总报表？openpyxl 新手实战指南",
  "description": "用 Python 自动整理 Excel 并生成汇总报表，核心流程是读取原始表格、清洗数据、按规则汇总、写入新工作表并设置格式。新手可优先使用 openpyxl 处理 xlsx 文件，它适合保留 Excel 样式、批量修改单元格和生成可直接交付的报表。",
  "keywords": [
    "Python 自动化",
    "Python Excel 自动化",
    "openpyxl 生成报表"
  ],
  "datePublished": "2026-06-02T01:15:24",
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
      "name": "openpyxl 和 pandas 处理 Excel 有什么区别？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "openpyxl 更适合操作 xlsx 文件本身，比如保留样式、合并单元格、写公式和调整列宽；pandas 更适合做数据分析和清洗。比如先用 pandas 汇总销售额，再用 openpyxl 生成带格式的报表。"
      }
    },
    {
      "@type": "Question",
      "name": "Python 可以自动整理多个 Excel 文件吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "可以，Python 能批量读取文件夹里的多个 Excel，再按统一规则清洗、合并和汇总。例如用 glob 找到 20 个 xlsx 文件，逐个读取“订单明细”工作表，最后生成一份月度汇总报表。"
      }
    },
    {
      "@type": "Question",
      "name": "openpyxl 能处理 xls 文件吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "openpyxl 不能直接处理 xls 文件，它主要支持 xlsx、xlsm 等较新的 Excel 格式。如果手里是 xls，可以先用 Excel 或 LibreOffice 转成 xlsx，再用 openpyxl 读取和整理。"
      }
    },
    {
      "@type": "Question",
      "name": "生成的 Excel 报表能保留公式吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "可以保留或写入公式，但 openpyxl 不负责计算公式结果。比如写入 =SUM(B2:B10) 后，通常需要用 Excel 打开文件时重新计算；读取时可用 data_only 参数选择读取公式或缓存值。"
      }
    },
    {
      "@type": "Question",
      "name": "这类 Python Excel 自动化能定时运行吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "可以定时运行，常见做法是把脚本放到服务器或本机任务计划中。Windows 可用“任务计划程序”每天 9 点执行，Linux 可用 cron 定时生成报表并保存到指定目录。"
      }
    }
  ]
}
</script>
<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "HowTo",
  "name": "怎么用 Python 自动整理 Excel 表格并生成汇总报表？openpyxl 新手实战指南",
  "description": "用 Python 自动整理 Excel 并生成汇总报表，核心流程是读取原始表格、清洗数据、按规则汇总、写入新工作表并设置格式。新手可优先使用 openpyxl 处理 xlsx 文件，它适合保留 Excel 样式、批量修改单元格和生成可直接交付的报表。",
  "step": [
    {
      "@type": "HowToStep",
      "name": "为什么适合用 Python 自动化整理 Excel 表格？",
      "text": "适合用 Python 自动化整理 Excel，因为它能把重复复制、筛选、求和、改格式变成可复用脚本，减少漏行、错公式和样式混乱。常见场景：销售日报、部门周报、库存统计、财务发票初筛；财务结果仍需人工复核，本节仅作技术说明。"
    },
    {
      "@type": "HowToStep",
      "name": "如何准备一个可运行的 Python Excel 自动化项目？",
      "text": "准备 Python Excel 自动化项目，只要先固定目录、安装 openpyxl（读写 xlsx 的 Python 库），再准备统一字段的示例表即可。建议只处理 xlsx（新版 Excel 工作簿格式），不直接处理旧版 xls，避免兼容问题。"
    },
    {
      "@type": "HowToStep",
      "name": "如何用 openpyxl 读取 Excel 并清洗原始数据？",
      "text": "用 openpyxl 读取 Excel 并清洗原始数据，做法是加载 Workbook（工作簿）、选中 Worksheet（工作表），从第 2 行跳过表头，把每行转成字典列表。清洗时重点处理空值、金额类型和字符串空格。"
    },
    {
      "@type": "HowToStep",
      "name": "如何按部门或产品汇总 Excel 数据并计算总额？",
      "text": "按部门或产品汇总 Excel 数据，就是遍历清洗后的明细行，按指定字段分组汇总（把相同部门、产品等归为一组），再累计订单数和销售额。"
    },
    {
      "@type": "HowToStep",
      "name": "如何生成带格式的 Excel 汇总报表？",
      "text": "生成带格式的 Excel 汇总报表，就是把清洗后的汇总结果写入新工作簿，添加标题、表头、样式和筛选后保存为 `summary_report.xlsx`。也可以在原工作簿中用 `workbook.create_sheet(\"部门汇总\")` 新建汇总工作表。"
    }
  ],
  "tool": [
    {
      "@type": "HowToTool",
      "name": "Python 3.10+、openpyxl 3.1+、Excel xlsx 文件"
    }
  ]
}
</script>
