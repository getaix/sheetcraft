# images

按单元格嵌入（in_cell）
- 通过工作簿 API：`ExcelWorkbook.insert_image_in_cell(sheet_name, cell, path, keep_ratio=True)`。
- 通过模板占位符：`{% img <path_expr> in_cell=true [keep_ratio=false] %}`。
- 行为说明：
  - `keep_ratio=true` 时在单元格宽/高边界内按比例缩放；`false` 时强制充满单元格（可能变形）。
  - openpyxl 下优先用 Pillow 读取原图尺寸；无 Pillow 时退化为近似值。
  - xlsxwriter 下依赖记录的行高/列宽（`set_row_height`/`set_column_width`），否则按默认估算。

::: sheetcraft.images