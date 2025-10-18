import os
import sys
import types

import pytest

from sheetcraft import ExcelWorkbook


def test_save_without_output_path_raises():
    wb = ExcelWorkbook(output_path=None)
    ws = wb.sheet("S")
    wb.write_row(ws, 1, ["A", 1])
    with pytest.raises(ValueError) as exc:
        wb.save()
    assert "Output path must be provided" in str(exc.value)


def test_parse_range_single_and_multi():
    wb = ExcelWorkbook()
    # 单个坐标
    r1, c1, r2, c2 = wb._parse_range("B7")
    assert (r1, c1, r2, c2) == (7, 2, 7, 2)
    # 范围坐标
    r1, c1, r2, c2 = wb._parse_range("A1:C3")
    assert (r1, c1, r2, c2) == (1, 1, 3, 3)


def test_get_sheet_default_for_xlsxwriter(monkeypatch, tmpfile):
    # 使用 xlsxwriter 桩，验证当未创建工作表时，get_sheet(None) 会创建默认工作表
    from tests.conftest import stub_xlsxwriter

    stub_xlsxwriter(monkeypatch)

    wb = ExcelWorkbook(output_path=tmpfile, fast=True)
    # 此时尚未创建任何工作表
    s = wb.get_sheet(None)
    # 应当创建并返回默认 "Sheet1"
    assert s is not None
    assert "Sheet1" in wb._sheets


def test_infer_format_unknown_extension():
    wb = ExcelWorkbook()
    fmt = wb._infer_format("file.zzz")
    assert fmt is None


def test_get_bytes_xlsxwriter_file_mode(monkeypatch, tmpfile):
    from tests.conftest import stub_xlsxwriter

    stub_xlsxwriter(monkeypatch)

    wb = ExcelWorkbook(output_path=tmpfile, fast=True)
    ws = wb.sheet("F")
    wb.write_row(ws, 1, ["X", 1])
    blob = wb.get_bytes()
    assert isinstance(blob, (bytes, bytearray))


def test_get_bytes_xlsxwriter_memory_mode(monkeypatch):
    from tests.conftest import stub_xlsxwriter

    stub_xlsxwriter(monkeypatch)
    # 兼容 xlsxwriter.Workbook(buffer, options) 的调用签名
    import xlsxwriter

    OrigWB = xlsxwriter.Workbook

    class WB(OrigWB):
        def __init__(self, path, options=None):
            super().__init__(path)

    monkeypatch.setattr(xlsxwriter, "Workbook", WB)

    wb = ExcelWorkbook(output_path=None, fast=True)
    ws = wb.sheet("M")
    wb.write_row(ws, 1, ["Y", 2])
    blob = wb.get_bytes()
    assert isinstance(blob, (bytes, bytearray))


def test_insert_image_in_cell_xlsxwriter_keep_ratio_false(monkeypatch, tmpfile):
    # 桩 xlsxwriter，设置行高/列宽，桩 PIL 返回固定尺寸，以验证缩放计算
    from tests.conftest import stub_xlsxwriter

    stub_xlsxwriter(monkeypatch)

    # 桩 PIL.Image.open，使其不读取真实文件
    pil = types.ModuleType("PIL")

    class FakeImage:
        def __init__(self, size):
            self.size = size
        def __enter__(self):
            return self
        def __exit__(self, exc_type, exc, tb):
            return False

    class Image:
        @staticmethod
        def open(path):
            return FakeImage((50, 100))

    pil.Image = Image
    monkeypatch.setitem(sys.modules, "PIL", pil)

    # 捕获插入调用以断言缩放参数
    import sheetcraft.workbook as fwb

    captured = {}

    def fake_insert(ws, row, col, path, scale_x=None, scale_y=None):
        captured["scale_x"] = scale_x
        captured["scale_y"] = scale_y

    monkeypatch.setattr(fwb, "insert_image_xlsxwriter", fake_insert)

    wb = ExcelWorkbook(output_path=tmpfile, fast=True)
    ws = wb.sheet("I")
    wb.set_column_width(ws, 2, 10.0)  # 列宽 -> 10*7+5 = 75px
    wb.set_row_height(ws, 2, 20.0)    # 行高 -> 20*96/72 ≈ 26.66px

    wb.insert_image_in_cell(ws, 2, 2, "dummy.jpg", keep_ratio=False)

    assert "scale_x" in captured and "scale_y" in captured
    # 期望：sx=75/50=1.5，sy≈26.67/100≈0.267
    assert pytest.approx(captured["scale_x"], rel=1e-3) == 1.5
    assert pytest.approx(captured["scale_y"], rel=1e-3) == int(round(20.0 * 96 / 72)) / 100