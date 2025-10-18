# fexcel

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

```bash
# 默认推荐：安装全部组件
pip install 'fexcel[all]'

# 或仅安装核心（不含可选组件）
pip install fexcel

# 分组件安装（可选）：
pip install 'fexcel[images]'      # 图片支持（Pillow）
pip install 'fexcel[xls]'         # 旧版 `.xls` 支持（xlwt/xlrd）
pip install 'fexcel[fast]'        # 更快的 `.xlsx` 写入（xlsxwriter）
pip install 'fexcel[template]'    # 模板渲染支持（Jinja2）
pip install 'fexcel[formula]'     # 公式评估（xlcalculator）
```

## Quick Start

```python
from fexcel import ExcelWorkbook

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
from fexcel.workbook import DataValidationSpec
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
  - Single-line form: place `{% for %}` on the row with templates and `{% endfor %}` on a later row; only the first item is rendered.
  - Multi-line block: put multiple template rows between `{% for %}` and `{% endfor %}`; capacity equals the number of template rows.

```python
from fexcel import ExcelTemplate, ExcelWorkbook

# Build a simple template that uses Jinja2 for (capacity=2)
wb = ExcelWorkbook(output_path='template.xlsx')
ws = wb.get_sheet()
wb.write_cell(ws, 1, 1, '{{ title }}')
wb.write_cell(ws, 3, 1, '{% for item in items %}')
wb.write_row(ws, 4, ['{{ item.name }}', '{{ item.qty }}', '{{ item.price }}', '=B4*C4'])
wb.write_row(ws, 5, ['{{ item.name }}', '{{ item.qty }}', '{{ item.price }}', '=B5*C5'])
wb.write_cell(ws, 6, 1, '{% endfor %}')
wb.save()

renderer = ExcelTemplate()
renderer.render('template.xlsx', {
    'title': 'Sales Report',
    'items': [
        {'name': 'Widget A', 'qty': 5, 'price': 19.99},
        {'name': 'Widget B', 'qty': 2, 'price': 29.50},
    ]
}, 'rendered.xlsx')
```

> 注意：旧版 `[[repeat ...]]` 语法已移除。请在模板中使用标准的 Jinja2 `{% for %}` 块。

## 格式修复模块（可选开关）

当生成的 `.xlsx` 在某些前端/预览组件中表现异常（如绘图锚点命名空间缺失），可以启用“格式修复”进行无损结构修补。

- 核心能力：`FormatFixConfig`（配置）、`FixReport`（结果报告）、`fix_xlsx(path, config)`（对文件进行修补）
- 目前内置规则：`prefix_drawing_anchors=True`，为绘图锚点标签增加 `xdr:` 前缀，并在根 `wsDr` 上补充 `xmlns:xdr` 命名空间。

启用方式：

- 在 `ExcelWorkbook` 层：
  - `apply_format_fix_on_save=True` 仅对 `.xlsx` 在 `save()` 后进行修补；`preview_temp()` 也会应用；`get_bytes()` 不做修改。
  - `format_fix_config=FormatFixConfig(...)` 配置具体规则。

```python
from fexcel import ExcelWorkbook, FormatFixConfig

wb = ExcelWorkbook(
    output_path='out.xlsx',
    fast=True,
    apply_format_fix_on_save=True,
    format_fix_config=FormatFixConfig(prefix_drawing_anchors=True)
)
ws = wb.add_sheet('Report')
# ... 写入数据 ...
wb.save()  # 保存后自动进行格式修复（仅 .xlsx）
```

- 在 `ExcelTemplate` 渲染层：
  - `apply_format_fix=True` 在渲染保存后对 `.xlsx` 进行修补。
  - `format_fix_config=FormatFixConfig(...)` 配置具体规则。

```python
from fexcel import ExcelTemplate, FormatFixConfig

renderer = ExcelTemplate(
    apply_format_fix=True,
    format_fix_config=FormatFixConfig(prefix_drawing_anchors=True)
)
renderer.render('template.xlsx', data, 'rendered.xlsx')
```

说明：修复过程只调整结构与命名空间，不改变单元格数据。若修复失败不会中断主流程（会记录报告或抛出详细异常用于排查）。

## Styles and Formatting

A simple style dict can include:

- `font`: `name`, `size`, `bold`, `italic`, `underline`, `color`
- `align`: `horizontal` (`left|center|right`), `vertical` (`top|center|bottom`), `wrap`
- `fill`: `color` (e.g., `#FFEECC`)
- `border`: `left|right|top|bottom` flags and optional `color`
- `number_format`: Excel number format string (e.g., `"#,##0.00"`, `"yyyy-mm-dd"`)

```python
price_style = {'number_format': '#,##0.00'}
wb.write_cell(ws, 2, 3, 19.99, price_style)
```

## Data Validation

```python
# Allow only values from a list
wb.add_data_validation(ws, 'E2:E100', DataValidationSpec(type='list', formula1='"Small,Medium,Large"'))

# Date between two bounds
wb.add_data_validation(ws, 'F2:F100', DataValidationSpec(type='date', operator='between', formula1='DATE(2024,1,1)', formula2='DATE(2024,12,31)'))
```

Note: `.xls` (xlwt) does not support data validation creation.

## Formula Support

Write formulas with `set_formula` or pass strings starting with `=` in `write_cell`.

Optional evaluation using xlcalculator:

```python
from fexcel.formulas import evaluate_xlsx_formulas
results = evaluate_xlsx_formulas('out.xlsx')
for (sheet, cell), value in results.items():
    print(sheet, cell, value)
```

## Multi-Sheet Operations

```python
wb = ExcelWorkbook(output_path='multi.xlsx', fast=False)
ws1 = wb.add_sheet('Summary')
ws2 = wb.add_sheet('Data')
wb.write_cell(ws1, 1, 1, 'Overview')
wb.write_rows(ws2, 1, [[1,2,3],[4,5,6],[7,8,9]])
wb.save()
```

## Legacy .xls Support

```python
wb = ExcelWorkbook(output_path='legacy.xls', file_format='xls')
ws = wb.add_sheet('Sheet1')
wb.write_rows(ws, 1, [["A", 1], ["B", 2]])
wb.insert_image(ws, 1, 3, 'logo.jpg')  # converted to BMP automatically
wb.save()
```

Constraints: `.xls` has limited styling, image, and no data validation creation.

## Performance for Large Datasets

- Use `fast=True` to write with xlsxwriter for very large `.xlsx` exports.
- For openpyxl, set `write_only=True` to stream rows efficiently.

```python
wb = ExcelWorkbook(output_path='big.xlsx', fast=True)
ws = wb.add_sheet('Data')
for i in range(2, 100_002):
    wb.write_row(ws, i, [f'Item {i}', i, i * 1.23])
wb.save()
```

## 技术实现细节概览

- 引擎选择：
  - `.xlsx` 默认使用 `openpyxl`；`fast=True` 时使用 `xlsxwriter`（更快写入，部分功能有限制）。
  - `.xls` 使用 `xlwt`（样式支持有限、无数据验证创建能力）。
  - `write_only=True` 可在 `openpyxl` 下流式写入以降低内存。
- 核心抽象：
  - `ExcelWorkbook` 提供工作表管理（`add_sheet/get_sheet`）、单元格与行写入（`write_cell/row/rows`）、行列尺寸、合并单元格、图片插入、数据验证等。
  - 样式通过统一字典进行描述，内部映射到不同引擎：`styles.py` 中的颜色规范化（`_normalize_color`）、`xlsxwriter_format_from_dict` 与 `xlwt_style_from_dict` 等。
- 模板渲染：
  - `ExcelTemplate.render` 基于 Jinja2，将单元格中的占位符与 `{% for %}` 块解析并渲染；多行块容量=模板行数，兼容常见导出场景。
- 格式修复：
  - `format_fix.py` 与生成流程解耦，仅在保存后对 `.xlsx` 进行结构性修补（如补全绘图锚点命名空间），保证数据不变。
- 公式评估：
  - `formulas.py` 提供 `evaluate_xlsx_formulas`，可选依赖 `xlcalculator` 进行公式计算并返回结果字典。
- 预览与字节：
  - `ExcelWorkbook.preview_temp()` 生成临时文件用于前端预览（若启用修复则也应用）；`get_bytes()` 直接返回内存字节，不做修复。
- 日志与错误：
  - 关键路径使用 `logging` 记录修复与渲染报告；异常均尽量包含上下文，便于定位问题。

## Project Structure

```
fexcel/
  __init__.py
  workbook.py
  styles.py
  images.py
  validation.py
  template.py
  formulas.py
examples/
  export_basic.py
  template_render.py
  image_insert.py
  multi_sheet.py
pyproject.toml
README.md
```

## 致谢

本项目感谢以下开源项目及其作者与维护者（致谢详情见 `NOTICE`）：
- openpyxl、XlsxWriter、xlwt、xlrd、Jinja2、Pillow、xlcalculator
- 开发工具：pytest、pytest-cov、pytest-benchmark

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

Repository: https://github.com/getaix/fexcel.git