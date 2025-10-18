import os
import sys
import types

import pytest

from sheetcraft import ExcelWorkbook


def test_insert_image_in_cell_openpyxl_without_pillow(monkeypatch, tmpfile):
    """验证在无 Pillow 的情况下，openpyxl 单元格内嵌图片路径可正常执行并生成文件。"""
    # 构造桩 Image 类，避免对 Pillow 的硬依赖
    fake_img_mod = types.SimpleNamespace()

    class FakeImage:
        def __init__(self, path):
            self.path = path
            self.width = 0
            self.height = 0

    fake_img_mod.Image = FakeImage
    fake_img_mod.PILImage = FakeImage
    monkeypatch.setitem(sys.modules, "openpyxl.drawing.image", fake_img_mod)

    wb = ExcelWorkbook(output_path=tmpfile)
    ws = wb.sheet("Embed")
    # 设置行高与列宽，触发尺寸计算分支
    wb.set_row_height(ws, 2, 24)
    wb.set_column_width(ws, 2, 18)
    # 使用不存在的图片路径以避免真实 I/O（桩路径不读取文件）
    wb.insert_image_in_cell(ws, 2, 2, "fake.png", keep_ratio=True)
    wb.save()
    assert os.path.exists(tmpfile)


def test_insert_image_in_cell_xlsxwriter_stub(monkeypatch, tmp_path):
    """验证 xlsxwriter 下的单元格内嵌图片缩放计算与保存流程（使用桩模块）。"""
    from tests.conftest import stub_xlsxwriter

    stub_xlsxwriter(monkeypatch)
    out = tmp_path / "embed_fast.xlsx"

    wb = ExcelWorkbook(output_path=str(out), fast=True)
    ws = wb.sheet("Fast")
    # 设置行高与列宽，以便计算缩放比例（即使无 Pillow 也应执行到调用层）
    wb.set_row_height(ws, 3, 26)
    wb.set_column_width(ws, 3, 20)
    # 使用不存在的图片路径；xlsxwriter 桩的 insert_image 为 no-op
    wb.insert_image_in_cell(ws, 3, 3, "fake.png", keep_ratio=True)
    wb.save()
    assert out.exists()


def test_insert_image_in_cell_xlsxwriter_no_ratio(monkeypatch, tmp_path):
    """验证 xlsxwriter 下在有 Pillow 时的非等比缩放路径（覆盖 keep_ratio=False 分支）。"""
    from tests.conftest import stub_xlsxwriter
    from PIL import Image

    stub_xlsxwriter(monkeypatch)
    out = tmp_path / "embed_fast_ratio_false.xlsx"
    png = tmp_path / "sized.png"
    # 生成一张已知尺寸的图片
    img = Image.new("RGB", (40, 20), (128, 128, 128))
    img.save(str(png), format="PNG")

    wb = ExcelWorkbook(output_path=str(out), fast=True)
    ws = wb.sheet("Fast2")
    wb.set_row_height(ws, 4, 40)
    wb.set_column_width(ws, 4, 18)
    wb.insert_image_in_cell(ws, 4, 4, str(png), keep_ratio=False)
    wb.save()
    assert out.exists()


def test_insert_image_in_cell_xlwt_stub(monkeypatch, tmp_path):
    """验证 .xls 引擎下的单元格内嵌图片路径（使用 xlwt 桩）。"""
    from tests.conftest import stub_xlwt

    stub_xlwt(monkeypatch)
    out = tmp_path / "embed_legacy.xls"

    wb = ExcelWorkbook(output_path=str(out), file_format="xls")
    ws = wb.sheet("Legacy")
    wb.set_row_height(ws, 2, 20)
    wb.set_column_width(ws, 2, 16)
    wb.insert_image_in_cell(ws, 2, 2, "fake.jpg", keep_ratio=True)
    wb.save()
    assert out.exists()