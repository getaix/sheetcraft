# 快速开始

## 安装

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

## 导出示例（.xlsx）

```python
from fexcel import ExcelWorkbook

wb = ExcelWorkbook(output_path='out.xlsx', fast=True)
ws = wb.add_sheet('Report')

# 样式示例
header_style = {
    'font': {'bold': True, 'size': 12},
    'fill': {'color': '#DDEEFF'},
    'align': {'horizontal': 'center'},
    'border': {'left': True, 'right': True, 'top': True, 'bottom': True}
}
wb.write_row(ws, 1, ['Item', 'Qty', 'Price', 'Total'], styles=[header_style]*4)

rows = [
    ['Widget A', 5, 19.99, '=B2*C2'],
    ['Widget B', 2, 29.50, '=B3*C3'],
]
wb.write_rows(ws, start_row=2, rows=rows)

from fexcel.workbook import DataValidationSpec
wb.add_data_validation(ws, 'B2:B100', DataValidationSpec(type='whole', operator='>=', formula1='0'))

wb.save()
```

## 模板渲染示例

```python
from fexcel import ExcelTemplate, ExcelWorkbook

# 构建简单模板（容量=2）
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

## 可选：启用格式修复

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
wb.save()
```

### 预览（可选）
仓库中提供了一个基于 Vue3 + Vite + TypeScript 的预览示例（`examples/preview`）。
- 支持选择或拖拽本地 `.xlsx` 文件进行预览。
- 前端会对部分关系路径进行轻量修复以提升图片兼容性。
- 启动：在 `examples/preview` 执行 `npm install && npm run dev`，打开浏览器访问本地地址即可。