import io
import os
import json
import tempfile
import pytest

from fexcel.workbook import ExcelWorkbook, DataValidationSpec


def test_data_validation_xlwt_noop():
    wb = ExcelWorkbook(output_path=None, file_format="xls")
    ws = wb.add_sheet("S")
    spec = DataValidationSpec(type="list", formula1="A1:A3")
    # xlwt 分支为 no-op，不应抛出异常
    wb.add_data_validation(ws, "A1:A3", spec)


def test_save_xlwt_success():
    fd, path = tempfile.mkstemp(suffix=".xls")
    os.close(fd)
    try:
        wb = ExcelWorkbook(output_path=path, file_format="xls")
        wb.add_sheet("S")
        # xlwt 分支调用 save，不应抛出异常
        wb.save(path)
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


def test_save_apply_format_fix_error_swallowed(monkeypatch):
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        wb = ExcelWorkbook(output_path=path, file_format="xlsx", apply_format_fix_on_save=True)
        ws = wb.add_sheet("S")
        wb.write_cell(ws, 1, 1, "x")

        # 让 fix_xlsx 抛错，save 应该吞掉异常
        import fexcel.format_fix as ff

        def boom(*args, **kwargs):
            raise RuntimeError("boom")

        monkeypatch.setattr(ff, "fix_xlsx", boom)
        wb.save(path)
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


def test_preview_temp_apply_format_fix_error_swallowed(monkeypatch):
    wb = ExcelWorkbook(output_path=None, file_format="xlsx", apply_format_fix_on_save=True)
    ws = wb.add_sheet("S")
    wb.write_cell(ws, 1, 1, "x")

    import fexcel.format_fix as ff

    def boom(*args, **kwargs):
        raise RuntimeError("boom")

    monkeypatch.setattr(ff, "fix_xlsx", boom)

    with wb.preview_temp() as path:
        assert os.path.exists(path)
        # 预览期间不应因格式修复失败而抛出异常


def test_export_dicts_empty_returns():
    wb = ExcelWorkbook(output_path=None, file_format="xlsx")
    ws = wb.add_sheet("S")
    # 空数据应直接返回，不抛异常
    wb.export_dicts(ws, 1, [], header_map=None, order=None)


def test_insert_image_in_cell_openpyxl_pil_fail(tmp_path, monkeypatch):
    # 使用不存在的路径触发 PIL 打开失败，从而覆盖异常分支
    missing = tmp_path / "no_such_image.png"
    wb = ExcelWorkbook(output_path=None, file_format="xlsx")
    ws = wb.add_sheet("S")
    # 强制 importlib 导入失败，使 openpyxl 分支走 fallback，避免实际读取文件
    import importlib

    def fail_import(name):
        raise ImportError("no openpyxl.drawing.image")

    monkeypatch.setattr(importlib, "import_module", fail_import)
    wb.insert_image_in_cell(ws, 1, 1, str(missing), keep_ratio=True)


def test_insert_image_in_cell_xlsxwriter_pil_fail(tmp_path, monkeypatch):
    missing = tmp_path / "no_such_image.png"
    wb = ExcelWorkbook(output_path=str(tmp_path / "a.xlsx"), file_format="xlsx", fast=True)
    ws = wb.add_sheet("S")
    # 给定列宽/行高以便计算缩放，但即使缺失图片也不应抛错
    wb.set_column_width(ws, 1, 12.0)
    wb.set_row_height(ws, 1, 20.0)
    # 在 workbook 模块替换调用，避免 xlsxwriter 实际读取文件
    import fexcel.workbook as fwb
    monkeypatch.setattr(fwb, "insert_image_xlsxwriter", lambda *args, **kwargs: None)
    wb.insert_image_in_cell(ws, 1, 1, str(missing), keep_ratio=True)