# sheetcraft

一个灵活、直观的 Python Excel 处理库。

- 支持 `.xlsx`（默认使用 openpyxl；可选 xlsxwriter 提升写入性能）
- 支持旧版 `.xls`（通过 xlwt，图片自动使用 BMP 回退）
- 高级导出能力：可定制单元格样式与格式
- 模板渲染（Jinja2），支持变量与块重复
- 图片插入（PNG/JPG；`.xls` 自动转 BMP）
- 数据有效性（Data Validation）辅助 API
- 公式书写与可选评估（xlcalculator）
- 多工作表操作
- 面向大数据集的性能优化

## Installation

> 环境要求：Python >= 3.11

```bash
# 推荐：安装常用组件（不含公式评估）
pip install 'sheetcraft[all]'

# 如需公式评估，请另外安装：
pip install xlcalculator

# 或仅安装核心（不含可选组件）
pip install sheetcraft

# 分组件安装（可选）：
pip install 'sheetcraft[images]'      # 图片支持（Pillow）
pip install 'sheetcraft[xls]'         # 旧版 `.xls` 支持（xlwt/xlrd）
pip install 'sheetcraft[fast]'        # 更快的 `.xlsx` 写入（xlsxwriter）
pip install 'sheetcraft[template]'    # 模板渲染支持（Jinja2）
pip install xlcalculator               # 公式评估（xlcalculator）
```

## Quick Start

```python
from sheetcraft import ExcelWorkbook

# Create a fast .xlsx export using xlsxwriter
wb = ExcelWorkbook(output_path='out.xlsx', fast=True)
ws = wb.add_sheet('Report')

# Write header with style
header_style = {
    'font': {'bold': True, 'size': 12},
    'fill': {'color': '#DDEEFF'},
    'align': {'horizontal': 'center'},
    'border': {'left': True, 'right': True, 'top': True, 'bottom': True}
}
wb.write_row(ws, 1, ['Item', 'Qty', 'Price', 'Total'], styles=[header_style]*4)

# Write data rows
rows = [
    ['Widget A', 5, 19.99, '=B2*C2'],
    ['Widget B', 2, 29.50, '=B3*C3'],
]
wb.write_rows(ws, start_row=2, rows=rows)

# Data validation for Qty column (whole number, >= 0)
from sheetcraft.workbook import DataValidationSpec
wb.add_data_validation(ws, 'B2:B100', DataValidationSpec(type='whole', operator='>=', formula1='0'))

# Insert an image
wb.insert_image(ws, 1, 6, 'logo.png', scale_x=0.5, scale_y=0.5)

# Set column widths
wb.set_column_width(ws, 1, 18)
wb.set_column_width(ws, 2, 10)
wb.set_column_width(ws, 3, 12)
wb.set_column_width(ws, 4, 12)

# Save
wb.save()
```

## Template Rendering (.xlsx)

Use Jinja2 placeholders in cell values and native `{% for %}` loops for repetition.

- `{{ title }}` replaces with `data['title']`
- `{% for item in items %}` ... `{% endfor %}` renders rows within the block capacity using `{{ item.field }}`.
- 单/多行块的容量由模板行数决定，超出时仅渲染容量对应的项。

...

## 兼容性与限制

- `.xlsx`：
  - `openpyxl` 功能更完整；适合多数样式与数据验证场景。
  - `xlsxwriter` 在 `fast=True` 时用于更快写入，部分高级样式/特性可能有限制；建议大数据量导出使用。
- `.xls`：
  - 样式能力有限、图片需 BMP 回退、不可创建数据验证；适合兼容旧系统的基础导出。
- 模板渲染：
  - 支持单元格占位符与 `{% for %}` 重复块；多行块的容量=模板行数。
  - 不支持跨工作表复杂引用的自动展开；如需更复杂逻辑建议用 Python 预处理数据。
- 公式评估：
  - 依赖 `xlcalculator`，仅对 `.xlsx` 提供可选评估；某些 Excel 特性（如外部引用、宏）不在支持范围内。
- 格式修复模块：
  - 仅对 `.xlsx` 在保存后进行结构修补；不改变单元格数据。
  - 当前主要适用绘图锚点命名空间修补；后续可按需扩展更多规则。

## License

Apache License 2.0。详见 `LICENSE`。

Repository: https://github.com/getaix/sheetcraft.git