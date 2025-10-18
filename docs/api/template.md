# ExcelTemplate

图片占位符（Jinja 扩展）
- 语法：`{% img <path_expr> [width=..] [height=..] [fit=true] [in_cell=true] [keep_ratio=false] %}`（首参为路径表达式，后续为命名参数；不使用逗号分隔）
- 参数：
  - `path_expr`：图片路径表达式，可为字符串或模板变量；相对路径按模板文件所在目录解析。
  - `width`/`height`：显式像素尺寸（openpyxl）。
  - `fit`：按单元格宽度适配（可能跨行），旧逻辑，优先级低于 `in_cell`。
  - `in_cell`：按单元格宽/高双向适配，尽量不跨单元格；默认配合 `keep_ratio`。
  - `keep_ratio`：与 `in_cell` 联用，是否保持原始宽高比，默认 `true`（`false` 时强制充满单元格）。
  - `scale_x`/`scale_y`：保留兼容（用于 xlsxwriter 缩放场景）。
- 行为：渲染阶段会扫描占位串并插入图片，插入成功后清空占位单元格的文本。
- 示例：
  - `"{% img 'examples/templates/img.png' in_cell=true %}"`

说明与兼容
- openpyxl 路径下在有 Pillow 时读取原图尺寸计算比例；无 Pillow 时回退按单元格近似尺寸插入。
- xlsxwriter 路径下缩放依赖已记录的行高/列宽（`set_row_height`/`set_column_width` 会被记录），否则按默认尺寸估算。

::: sheetcraft.template.ExcelTemplate