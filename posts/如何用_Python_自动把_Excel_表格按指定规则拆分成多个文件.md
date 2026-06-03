---
title: 如何用 Python 自动把 Excel 表格按指定规则拆分成多个文件？
author: 天机枢
source_question: 如何用 Python 自动把 Excel 表格按指定规则拆分成多个文件？
intent: how-to
difficulty: beginner
target_keywords:
- Python 自动化
- Python 拆分 Excel
- 办公自动化脚本
description: 用 Python 拆分 Excel 的核心做法是：先用 pandas 读取表格，再按指定列、行数或条件分组，最后循环导出为多个独立文件。相比手动复制粘贴，办公自动化脚本更适合处理部门报表、客户清单、地区数据等重复拆分任务，初学者只需掌握
  read_excel、groupby 和 to_excel 即
environment: Python 3.10+、pandas 2.x、openpyxl 3.x、Excel xlsx 文件
generated_at: '2026-06-03T09:41:23'
updated: '2026-06-03'
generator: opc-geo v0.0.1
---


# 如何用 Python 自动把 Excel 表格按指定规则拆分成多个文件？


*📅 最后更新:2026-06-03 · 🛠 运行环境:Python 3.10+、pandas 2.x、openpyxl 3.x、Excel xlsx 文件*


> **TL;DR**:用 Python 拆分 Excel 的核心做法是：先用 pandas 读取表格，再按指定列、行数或条件分组，最后循环导出为多个独立文件。相比手动复制粘贴，办公自动化脚本更适合处理部门报表、客户清单、地区数据等重复拆分任务，初学者只需掌握 read_excel、groupby 和 to_excel 即可完成。


## 为什么适合用 Python 自动拆分 Excel 文件？

适合用 Python 自动拆分 Excel，因为规则固定、重复多，脚本比人工更快、更准、可复用。典型场景：按部门拆月报、按地区拆订单、按销售人员拆业绩、按客户类型拆名单。

### 对比卡片

- **手动复制粘贴**：适合 1-2 份；优点：零门槛；缺点：慢，易漏行错列。
- **Excel 筛选另存**：适合少量分类；优点：直观；缺点：每次重复操作。
- **Python 脚本**：适合几十到上百组；优点：一次编写反复运行；缺点：需配置环境。

本文用 pandas（Python 表格数据分析库），不只用 openpyxl（Excel 文件读写库）：pandas 的 read_excel、groupby、筛选更适合表格分组拆分。

## 如何准备 Python 拆分 Excel 所需的环境和示例文件？

准备 Python 拆分 Excel，先安装依赖、确认源文件字段，并固定路径、工作表名和输出目录。安装命令：`pip install pandas openpyxl`，pandas（数据分析库）负责读写和分组，openpyxl（Excel 引擎库）负责处理 `.xlsx`。

准备清单：
- Python 3.9+
- 依赖库：pandas、openpyxl
- 源 Excel 文件：`sample.xlsx`
- 拆分字段：如 `部门`、`地区`
- 输出文件夹：如 `output_files/`

示例表可包含 `姓名`、`部门`、`地区`、`销售额`。文件路径、工作表名称、列名必须与代码一致；输出目录用于集中保存拆分结果，避免混在原目录。

这段代码会创建后续可直接使用的 `sample.xlsx`：

```python
from pathlib import Path
import pandas as pd

output_dir = Path("output_files")
output_dir.mkdir(exist_ok=True)  # 创建输出目录，用于保存拆分后的文件

df = pd.DataFrame(
    [
        {"姓名": "张三", "部门": "销售部", "地区": "华东", "销售额": 12000},
        {"姓名": "李四", "部门": "销售部", "地区": "华南", "销售额": 9800},
        {"姓名": "王五", "部门": "运营部", "地区": "华东", "销售额": 7600},
        {"姓名": "赵六", "部门": "财务部", "地区": "华北", "销售额": 6500},
        {"姓名": "钱七", "部门": "运营部", "地区": "华南", "销售额": 8300},
    ]
)

workbook = "sample.xlsx"
df.to_excel(workbook, sheet_name="明细", index=False)  # 写入名为“明细”的工作表
```

运行后，当前目录会生成 `sample.xlsx` 和 `output_files` 文件夹。

## 如何按某一列的值把 Excel 拆分成多个文件？

按某一列拆分 Excel，就是用 `read_excel` 读取数据，再用 `groupby`（按字段把数据分组）按“部门、地区、负责人”等列分组，最后把每组用 `to_excel` 导出成独立文件。文件名要过滤 `/ \ : * ? " < > |` 等特殊字符，避免保存失败。

| 函数 | 作用 | 初学者注意 |
|---|---|---|
| `read_excel` | 读取 xlsx | 需安装 `openpyxl` |
| `groupby` | 按列分组 | 列名必须完全一致 |
| `to_excel` | 导出文件 | 建议 `index=False` |
| `os.makedirs` | 创建文件夹 | 加 `exist_ok=True` |

这段代码会读取 `sample.xlsx`，按“部门”列拆分，并保存到 `output_by_department` 文件夹：

```python
import os
import re
import pandas as pd

input_file = "sample.xlsx"
output_dir = "output_by_department"
group_column = "部门"

# 如果没有示例文件，自动创建一个，方便直接运行
if not os.path.exists(input_file):
    sample_data = {
        "部门": ["销售", "销售", "财务", "技术", "技术"],
        "姓名": ["张三", "李四", "王五", "赵六", "孙七"],
        "金额": [1200, 1800, 900, 2600, 3100],
    }
    pd.DataFrame(sample_data).to_excel(input_file, index=False)

# 创建输出文件夹
os.makedirs(output_dir, exist_ok=True)

# 读取 Excel 文件
df = pd.read_excel(input_file)

# 清理文件名中的非法字符
def clean_filename(name):
    return re.sub(r'[\\/:*?"<>|]', "_", str(name)).strip()

# 按“部门”列分组并分别导出
for department, group_df in df.groupby(group_column):
    safe_department = clean_filename(department)
    output_file = os.path.join(output_dir, f"部门_{safe_department}.xlsx")
    group_df.to_excel(output_file, index=False)  # 不导出 pandas 默认索引列

print("拆分完成")
```

运行后会生成类似 `部门_销售.xlsx`、`部门_财务.xlsx`、`部门_技术.xlsx` 的文件。

## 如何按固定行数把 Excel 平均拆分成多个文件？

按固定行数拆分 Excel，就是把 DataFrame（pandas 的二维表格对象）按每 500 行、1000 行或指定数量切成多段，再分别导出。它适合大表分页、系统批量导入、单个文件限制行数的场景。流程如下：

- 读取文件
- 设定每份行数
- 循环切片
- 导出文件

核心逻辑是用 `range(0, 总行数, 每份行数)` 分段截取，并用 `part_001.xlsx`、`part_002.xlsx` 命名。`to_excel` 默认会在每个拆分文件中保留表头。

这段代码会把 `sample.xlsx` 按每 100 行拆分到 `output_by_rows` 目录：

```python
from pathlib import Path
import pandas as pd

source_file = Path("sample.xlsx")
output_dir = Path("output_by_rows")
rows_per_file = 100

# 如果 sample.xlsx 不存在，创建一个示例文件，方便直接运行测试
if not source_file.exists():
    sample_data = {
        "姓名": [f"用户{i}" for i in range(1, 251)],
        "部门": ["销售部", "技术部", "财务部", "运营部", "客服部"] * 50,
        "金额": [i * 10 for i in range(1, 251)]
    }
    sample_df = pd.DataFrame(sample_data)
    sample_df.to_excel(source_file, index=False)

# 创建输出目录
output_dir.mkdir(exist_ok=True)

# 读取 Excel 文件
df = pd.read_excel(source_file)

# 按固定行数循环切片并导出
for start_row in range(0, len(df), rows_per_file):
    part_number = start_row // rows_per_file + 1
    part_df = df.iloc[start_row:start_row + rows_per_file]

    output_file = output_dir / f"part_{part_number:03d}.xlsx"

    # index=False 不导出行号；表头会默认保留
    part_df.to_excel(output_file, index=False)

print(f"拆分完成，共生成 {(len(df) + rows_per_file - 1) // rows_per_file} 个文件。")
```

运行后会生成 `part_001.xlsx`、`part_002.xlsx` 等文件，每个文件最多 100 行数据。

## 如何让拆分脚本更稳定并避免常见错误？

让拆分脚本更稳定的关键，是在读取前查路径、处理前查列名、导出前处理空值和文件名。中文路径通常可用，但建议用 `pathlib`（Python 的路径对象工具）统一管理路径，减少斜杠和转义错误。

| 常见错误 | 原因 | 解决方法 |
|---|---|---|
| `FileNotFoundError` | 源 Excel 不存在 | 用 `Path.exists()` 检查 |
| `KeyError` | 拆分列名写错 | 打印 `df.columns` 核对 |
| 缺少 `openpyxl` | 未安装 Excel 读写引擎 | 执行 `pip install openpyxl` |
| `PermissionError` | 文件被 Excel 占用 | 关闭文件后重跑 |
| 路径不存在 | 输出目录未创建 | 用 `mkdir()` 自动创建 |

下面代码按“配置区、读取区、处理区、导出区”组织，可直接运行并按指定列拆分 Excel：

```python
from pathlib import Path
import re
import pandas as pd

# ===== 配置区 =====
source_file = Path("demo_split.xlsx")
output_dir = Path("split_output")
split_column = "部门"
create_demo = True

def safe_filename(value):
    """把分组值转换成安全文件名"""
    if pd.isna(value) or str(value).strip() == "":
        text = "空值"
    else:
        text = str(value).strip()
    text = re.sub(r'[\\/:*?"<>|]+', "_", text)
    return text[:80] or "未命名"

# 生成最小示例文件，便于复制后直接运行
if create_demo and not source_file.exists():
    demo_df = pd.DataFrame({
        "部门": ["销售部", "技术部", "销售部", "财务部", None],
        "姓名": ["张三", "李四", "王五", "赵六", "钱七"],
        "金额": [100, 200, 150, 300, 80]
    })
    demo_df.to_excel(source_file, index=False, engine="openpyxl")

# ===== 读取区 =====
if not source_file.exists():
    raise FileNotFoundError(f"源文件不存在：{source_file.resolve()}")

try:
    df = pd.read_excel(source_file, engine="openpyxl")
except ImportError as exc:
    raise ImportError("缺少 openpyxl，请先执行：pip install openpyxl") from exc

if df.empty:
    raise ValueError("源 Excel 没有数据，无法拆分")

if split_column not in df.columns:
    raise KeyError(f"找不到拆分列：{split_column}，当前列名：{list(df.columns)}")

# ===== 处理区 =====
output_dir.mkdir(parents=True, exist_ok=True)
used_names = {}
export_count = 0

# ===== 导出区 =====
for group_value, group_df in df.groupby(split_column, dropna=False):
    if group_df.empty:
        continue

    base_name = safe_filename(group_value)
    used_names[base_name] = used_names.get(base_name, 0) + 1

    if used_names[base_name] == 1:
        file_name = f"{base_name}.xlsx"
    else:
        file_name = f"{base_name}_{used_names[base_name]}.xlsx"

    target_file = output_dir / file_name

    try:
        group_df.to_excel(target_file, index=False, engine="openpyxl")
    except PermissionError as exc:
        raise PermissionError(f"文件可能被 Excel 占用，请关闭后重试：{target_file}") from exc

    export_count += 1
    print(f"已导出：{target_file.resolve()}，行数：{len(group_df)}")

print(f"拆分完成，共导出 {export_count} 个文件")
```

运行后，终端会打印每个导出文件的路径和行数，结果文件保存在 `split_output` 目录。

## 要点回顾

- Python 拆分 Excel 最常用的方法是 pandas 读取数据、按规则分组或切片、再循环导出文件。
- 按列拆分适合部门、地区、客户等分类场景，按固定行数拆分适合大文件分页和批量导入场景。
- 初学者应重点检查文件路径、工作表名称、拆分列名和输出目录，避免脚本运行时报错。

## 常见问答(FAQ)

**Q:Python 拆分 Excel 必须安装 Excel 软件吗？**

A:不必须安装 Excel 软件。Python 通过 pandas、openpyxl 等库直接读写 .xlsx 文件，服务器或没有 Office 的电脑也能运行。只要文件格式支持，脚本就可以读取、分组并导出。

**Q:可以按多个条件拆分 Excel 吗？**

A:可以按多个条件拆分。常见做法是用 pandas 按多列 groupby，例如同时按“部门”和“地区”分组，把销售部-华东、销售部-华南分别导出成不同文件。

**Q:拆分后能保留原来的格式吗？**

A:默认用 pandas 导出时，通常只能保留数据，原来的颜色、合并单元格、列宽等格式可能丢失。若需要保留格式，可结合 openpyxl 复制样式，但代码会更复杂。

**Q:Excel 文件很大时运行很慢怎么办？**

A:文件很大时应减少一次性处理的数据量。可以只读取需要的列、按条件过滤后再拆分，或把 20 万行数据分批处理；如果只是数据表，CSV 往往比 Excel 更快。

**Q:为什么代码提示找不到 openpyxl？**

A:通常是因为当前 Python 环境没有安装 openpyxl，或安装到了另一个环境。可先运行 pip install openpyxl，再确认编辑器、命令行和脚本使用的是同一个 Python 解释器。

<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "Article",
  "headline": "如何用 Python 自动把 Excel 表格按指定规则拆分成多个文件？",
  "description": "用 Python 拆分 Excel 的核心做法是：先用 pandas 读取表格，再按指定列、行数或条件分组，最后循环导出为多个独立文件。相比手动复制粘贴，办公自动化脚本更适合处理部门报表、客户清单、地区数据等重复拆分任务，初学者只需掌握 read_excel、groupby 和 to_excel 即",
  "keywords": [
    "Python 自动化",
    "Python 拆分 Excel",
    "办公自动化脚本"
  ],
  "datePublished": "2026-06-03T09:41:23",
  "dateModified": "2026-06-03",
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
      "name": "Python 拆分 Excel 必须安装 Excel 软件吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "不必须安装 Excel 软件。Python 通过 pandas、openpyxl 等库直接读写 .xlsx 文件，服务器或没有 Office 的电脑也能运行。只要文件格式支持，脚本就可以读取、分组并导出。"
      }
    },
    {
      "@type": "Question",
      "name": "可以按多个条件拆分 Excel 吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "可以按多个条件拆分。常见做法是用 pandas 按多列 groupby，例如同时按“部门”和“地区”分组，把销售部-华东、销售部-华南分别导出成不同文件。"
      }
    },
    {
      "@type": "Question",
      "name": "拆分后能保留原来的格式吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "默认用 pandas 导出时，通常只能保留数据，原来的颜色、合并单元格、列宽等格式可能丢失。若需要保留格式，可结合 openpyxl 复制样式，但代码会更复杂。"
      }
    },
    {
      "@type": "Question",
      "name": "Excel 文件很大时运行很慢怎么办？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "文件很大时应减少一次性处理的数据量。可以只读取需要的列、按条件过滤后再拆分，或把 20 万行数据分批处理；如果只是数据表，CSV 往往比 Excel 更快。"
      }
    },
    {
      "@type": "Question",
      "name": "为什么代码提示找不到 openpyxl？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "通常是因为当前 Python 环境没有安装 openpyxl，或安装到了另一个环境。可先运行 pip install openpyxl，再确认编辑器、命令行和脚本使用的是同一个 Python 解释器。"
      }
    }
  ]
}
</script>
<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "HowTo",
  "name": "如何用 Python 自动把 Excel 表格按指定规则拆分成多个文件？",
  "description": "用 Python 拆分 Excel 的核心做法是：先用 pandas 读取表格，再按指定列、行数或条件分组，最后循环导出为多个独立文件。相比手动复制粘贴，办公自动化脚本更适合处理部门报表、客户清单、地区数据等重复拆分任务，初学者只需掌握 read_excel、groupby 和 to_excel 即",
  "step": [
    {
      "@type": "HowToStep",
      "name": "为什么适合用 Python 自动拆分 Excel 文件？",
      "text": "适合用 Python 自动拆分 Excel，因为规则固定、重复多，脚本比人工更快、更准、可复用。典型场景：按部门拆月报、按地区拆订单、按销售人员拆业绩、按客户类型拆名单。"
    },
    {
      "@type": "HowToStep",
      "name": "如何准备 Python 拆分 Excel 所需的环境和示例文件？",
      "text": "准备 Python 拆分 Excel，先安装依赖、确认源文件字段，并固定路径、工作表名和输出目录。安装命令：`pip install pandas openpyxl`，pandas（数据分析库）负责读写和分组，openpyxl（Excel 引擎库）负责处理 `.xlsx`。"
    },
    {
      "@type": "HowToStep",
      "name": "如何按某一列的值把 Excel 拆分成多个文件？",
      "text": "按某一列拆分 Excel，就是用 `read_excel` 读取数据，再用 `groupby`（按字段把数据分组）按“部门、地区、负责人”等列分组，最后把每组用 `to_excel` 导出成独立文件。文件名要过滤 `/ \\ : * ? \" < > |` 等特殊字符，避免保存失败。"
    },
    {
      "@type": "HowToStep",
      "name": "如何按固定行数把 Excel 平均拆分成多个文件？",
      "text": "按固定行数拆分 Excel，就是把 DataFrame（pandas 的二维表格对象）按每 500 行、1000 行或指定数量切成多段，再分别导出。它适合大表分页、系统批量导入、单个文件限制行数的场景。流程如下："
    },
    {
      "@type": "HowToStep",
      "name": "如何让拆分脚本更稳定并避免常见错误？",
      "text": "让拆分脚本更稳定的关键，是在读取前查路径、处理前查列名、导出前处理空值和文件名。中文路径通常可用，但建议用 `pathlib`（Python 的路径对象工具）统一管理路径，减少斜杠和转义错误。"
    }
  ],
  "tool": [
    {
      "@type": "HowToTool",
      "name": "Python 3.10+、pandas 2.x、openpyxl 3.x、Excel xlsx 文件"
    }
  ]
}
</script>
