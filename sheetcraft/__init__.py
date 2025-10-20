"""
sheetcraft：灵活且直观的 Excel 处理库。

提供高层 API：数据导出、单元格样式、图片插入、模板渲染、
数据有效性、公式处理以及多工作表操作。

支持 `.xlsx`（默认使用 openpyxl，可选 xlsxwriter 以提升写入性能），
以及旧版 `.xls`（通过 xlwt，图片使用 BMP 回退以兼容）。
"""

from .workbook import ExcelWorkbook
from .template import ExcelTemplate
from .format_fix import FormatFixConfig, FixReport, fix_xlsx

__all__ = [
    "ExcelWorkbook",
    "ExcelTemplate",
    "FormatFixConfig",
    "FixReport",
    "fix_xlsx",
]

__version__ = "0.2.0"
