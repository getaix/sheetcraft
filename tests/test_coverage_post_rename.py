import types
import pytest

from sheetcraft import ExcelWorkbook
import os


def test_save_without_output_path_raises():
    wb = ExcelWorkbook()
    with pytest.raises(ValueError):
        wb.save()


def test_insert_image_in_cell_openpyxl_no_ratio(monkeypatch, tmp_path):
    # 记录传入的宽高以断言 keep_ratio=False 分支
    captured = {}

    def fake_insert(ws, r, c, path, width=None, height=None):
        captured["args"] = (r, c, path)
        captured["size"] = (width, height)

    monkeypatch.setattr("sheetcraft.workbook.insert_image_openpyxl", fake_insert)

    out = tmp_path / "img_openpyxl.xlsx"
    wb = ExcelWorkbook(output_path=str(out), fast=False)
    ws = wb.sheet("S")
    # 设置行高列宽以触发尺寸计算
    wb.set_row_height(ws, 2, 20.0)   # 20pt -> ≈ 26 px
    wb.set_column_width(ws, 2, 10.0) # 10 char -> ≈ 75 px

    wb.insert_image_in_cell(ws, 2, 2, "fake.png", keep_ratio=False)
    assert captured.get("args") == (2, 2, "fake.png")
    w, h = captured.get("size", (0, 0))
    # 近似检查：宽度约 75px，高度约 26~27px
    assert abs(w - 75) <= 3
    assert 24 <= h <= 28


def test_insert_image_in_cell_xlsxwriter_no_ratio(monkeypatch, tmp_path):
    # 使用 conftest 的 xlsxwriter 桩，避免真实依赖
    from tests.conftest import stub_xlsxwriter
    import sys
    
    stub_xlsxwriter(monkeypatch)

    # 注入简单的 PIL 桩以提供图像尺寸
    class _Img:
        def __init__(self):
            self.size = (100, 100)
        def __enter__(self):
            return self
        def __exit__(self, exc_type, exc, tb):
            return False
    class _PIL:
        class Image:
            @staticmethod
            def open(path):
                return _Img()
    monkeypatch.setitem(sys.modules, "PIL", _PIL)

    out = tmp_path / "img_xlsxwriter.xlsx"
    wb = ExcelWorkbook(output_path=str(out), fast=True)
    ws = wb.add_sheet("S1")

    # 设置行高列宽，供缩放计算
    wb.set_row_height(ws, 3, 30.0)   # 30pt -> ≈ 40 px
    wb.set_column_width(ws, 3, 15.0) # 15 char -> ≈ 110 px

    # 替换插入函数以记录缩放参数
    last = {}

    def fake_insert(ws2, r, c, path, scale_x=None, scale_y=None):
        last["args"] = (r, c, path)
        last["scale"] = (scale_x, scale_y)

    monkeypatch.setattr("sheetcraft.workbook.insert_image_xlsxwriter", fake_insert)

    wb.insert_image_in_cell(ws, 3, 3, "fake.png", keep_ratio=False)
    assert last.get("args") == (3, 3, "fake.png")
    sx, sy = last.get("scale", (None, None))
    assert sx is not None and sy is not None
    assert sx > 0 and sy > 0


def test_insert_image_in_cell_xlwt_branch(monkeypatch, tmp_path):
    # 使用 xlwt 桩，验证 xlwt 分支路径被执行
    from tests.conftest import stub_xlwt

    stub_xlwt(monkeypatch)
    out = tmp_path / "img_xlwt.xls"
    wb = ExcelWorkbook(output_path=str(out), file_format="xls")
    ws = wb.add_sheet("S2")

    called = {"v": False}

    def fake_insert(ws2, r, c, path):
        called["v"] = True

    monkeypatch.setattr("sheetcraft.workbook.insert_image_xlwt", fake_insert)

    wb.insert_image_in_cell(ws, 2, 2, "fake.png")
    assert called["v"] is True


def test_insert_image_in_cell_xlsxwriter_keep_ratio(monkeypatch, tmp_path):
    from tests.conftest import stub_xlsxwriter
    import sys

    stub_xlsxwriter(monkeypatch)

    class _Img:
        def __init__(self):
            self.size = (200, 100)
        def __enter__(self):
            return self
        def __exit__(self, exc_type, exc, tb):
            return False
    class _PIL:
        class Image:
            @staticmethod
            def open(path):
                return _Img()
    monkeypatch.setitem(sys.modules, "PIL", _PIL)

    out = tmp_path / "img_ratio.xlsx"
    wb = ExcelWorkbook(output_path=str(out), fast=True)
    ws = wb.add_sheet("R")
    wb.set_row_height(ws, 4, 24.0)
    wb.set_column_width(ws, 4, 12.0)

    last = {}
    def fake_insert(ws2, r, c, path, scale_x=None, scale_y=None):
        last["args"] = (r, c, path)
        last["scale"] = (scale_x, scale_y)
    monkeypatch.setattr("sheetcraft.workbook.insert_image_xlsxwriter", fake_insert)

    wb.insert_image_in_cell(ws, 4, 4, "fake.png", keep_ratio=True)
    sx, sy = last.get("scale", (None, None))
    assert sx is not None and sy is not None
    assert abs(sx - sy) < 1e-6


def test_insert_image_in_cell_openpyxl_keep_ratio(monkeypatch, tmp_path):
    import sys
    captured = {}

    class _Img:
        def __init__(self):
            self.size = (200, 100)
        def __enter__(self):
            return self
        def __exit__(self, exc_type, exc, tb):
            return False
    class _PIL:
        class Image:
            @staticmethod
            def open(path):
                return _Img()
    monkeypatch.setitem(sys.modules, "PIL", _PIL)

    def fake_insert(ws2, r, c, path, width=None, height=None):
        captured["args"] = (r, c, path)
        captured["size"] = (width, height)
    monkeypatch.setattr("sheetcraft.workbook.insert_image_openpyxl", fake_insert)

    out = tmp_path / "img_openpyxl_ratio.xlsx"
    wb = ExcelWorkbook(output_path=str(out), fast=False)
    ws = wb.sheet("O")
    wb.set_row_height(ws, 2, 20.0)
    wb.set_column_width(ws, 2, 10.0)

    wb.insert_image_in_cell(ws, 2, 2, "fake.png", keep_ratio=True)
    w, h = captured.get("size", (0, 0))
    assert w > 0 and h > 0
    # 保持比例时，尺寸应不超过单元格尺寸
    assert w <= 80 and h <= 30


def test_get_sheet_default_creation_for_xlsxwriter(monkeypatch, tmp_path):
    # 确认在无工作表时，get_sheet() 会创建默认表（xlsxwriter）
    from tests.conftest import stub_xlsxwriter

    stub_xlsxwriter(monkeypatch)
    out = tmp_path / "default.xlsx"
    wb = ExcelWorkbook(output_path=str(out), fast=True)

    sheet = wb.get_sheet()
    # 由于 xlsxwriter 无 active sheet，应创建默认 Sheet1
    assert sheet is not None
    # 保存以确保工作簿生命周期正确
    wb.save()
    assert os.path.exists(out)