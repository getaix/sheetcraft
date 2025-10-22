# 快速开始

> 环境要求：Python >= 3.11

## 安装

```bash
# 常用组件（不含公式评估）
pip install 'sheetcraft[all]'

# 如需公式评估，请另外安装：
pip install xlcalculator
```

## 导出示例（.xlsx）

```python
from sheetcraft import ExcelWorkbook

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

from sheetcraft.workbook import DataValidationSpec
wb.add_data_validation(ws, 'B2:B100', DataValidationSpec(type='whole', operator='>=', formula1='0'))

wb.save()
```

## 模板渲染示例

```python
from sheetcraft.template import ExcelTemplate

renderer = ExcelTemplate()
# 具体模板和数据示例请参考 README 与 docs/api
```

## 可选：启用格式修复

```python
from sheetcraft import ExcelWorkbook, FormatFixConfig

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