import os
import sys
import types

from fexcel import ExcelWorkbook
from fexcel.images import insert_image_xlsxwriter


def test_openpyxl_insert_image_without_pillow(monkeypatch, tmpfile):
    """通过桩替换 openpyxl Image，验证图片插入路径不依赖 Pillow。"""
    # 构造桩 Image 类
    fake_img_mod = types.SimpleNamespace()

    class FakeImage:
        def __init__(self, path):
            self.path = path
            self.width = 0
            self.height = 0

    fake_img_mod.Image = FakeImage
    # openpyxl.reader.drawings 也会从该模块导入 PILImage，提供桩以避免依赖 Pillow
    fake_img_mod.PILImage = FakeImage
    monkeypatch.setitem(sys.modules, "openpyxl.drawing.image", fake_img_mod)

    wb = ExcelWorkbook(output_path=tmpfile)
    ws = wb.sheet("I")
    wb.insert_image(ws, 2, 2, "fake.png", width=100, height=80)
    wb.save()
    assert os.path.exists(tmpfile)


def test_xlsxwriter_insert_image_options():
    """验证 xlsxwriter 图片插入选项传递。"""

    class WS:
        def __init__(self):
            self.last = None

        def insert_image(self, r, c, path, opts=None):
            self.last = (r, c, path, opts)

    ws = WS()
    insert_image_xlsxwriter(ws, 3, 4, "x.png", scale_x=0.5, scale_y=0.6)
    assert ws.last and ws.last[3]["x_scale"] == 0.5 and ws.last[3]["y_scale"] == 0.6


import types
from fexcel.images import (
    insert_image_openpyxl,
    insert_image_xlwt,
    calc_fit_size_openpyxl,
)


def test_insert_image_openpyxl_fallback(monkeypatch):
    # 强制 fallback：模块存在但 __spec__ 为 None
    fake_img_mod = types.SimpleNamespace(Image=None, __spec__=None)
    monkeypatch.setitem(sys.modules, "openpyxl.drawing.image", fake_img_mod)

    class WS:
        def __init__(self):
            self.column_dimensions = {}
            self.row_dimensions = {}
            self.calls = []

        def add_image(self, img, anchor):
            self.calls.append((img, anchor))

    ws = WS()
    insert_image_openpyxl(ws, 1, 1, "fake.png", width=50, height=40)
    # fallback 情况下不进行插入，以避免对 Pillow/openpyxl 的硬依赖
    assert ws.calls == []


def test_insert_image_openpyxl_two_cell_anchor(monkeypatch):
    # 非 fallback：提供完整桩模块
    class StubImage:
        def __init__(self, path):
            self.path = path
            self.width = 0
            self.height = 0

    fake_img_mod = types.SimpleNamespace(Image=StubImage)
    fake_img_mod.__spec__ = object()
    monkeypatch.setitem(sys.modules, "openpyxl.drawing.image", fake_img_mod)

    class AnchorMarker:
        def __init__(self, col, colOff, row, rowOff):
            self.col = col
            self.row = row

    class TwoCellAnchor:
        def __init__(self, _from, to):
            self._from = _from
            self.to = to

    fake_sp_mod = types.SimpleNamespace(AnchorMarker=AnchorMarker, TwoCellAnchor=TwoCellAnchor)
    monkeypatch.setitem(sys.modules, "openpyxl.drawing.spreadsheet_drawing", fake_sp_mod)

    class Dim:  # 提供宽高
        def __init__(self, width=None, height=None):
            self.width = width
            self.height = height

    class WS:
        def __init__(self):
            self.column_dimensions = {"B": Dim(width=10)}
            self.row_dimensions = {2: Dim(height=20.0)}
            self.calls = []

        def add_image(self, img, anchor):
            self.calls.append(anchor)

    ws = WS()
    insert_image_openpyxl(ws, 2, 2, "fake.png", width=100, height=80)
    assert ws.calls and isinstance(ws.calls[0], TwoCellAnchor)


def test_insert_image_xlwt_and_cleanup(monkeypatch, tmp_path):
    # 伪造 BMP 转换与插入，并验证临时文件被删除
    bmp_file = tmp_path / "tmp.bmp"

    def fake_convert(path):
        bmp_file.write_bytes(b"\x00")
        return str(bmp_file)

    monkeypatch.setattr("fexcel.images.convert_image_to_bmp_temp", fake_convert)

    class WS:
        def __init__(self):
            self.called = False

        def insert_bitmap(self, p, r, c):
            self.called = True

    ws = WS()
    insert_image_xlwt(ws, 3, 4, "src.png")
    assert ws.called is True
    assert not bmp_file.exists()


def test_calc_fit_size_openpyxl_with_pillow(monkeypatch):
    # 注入 PIL.Image 桩，验证按宽度充满比例缩放
    class StubCtx:
        def __init__(self):
            self.size = (200, 100)

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb):
            return False

    class StubImage:
        @staticmethod
        def open(path):
            return StubCtx()

    fake_pil = types.SimpleNamespace(Image=StubImage)
    monkeypatch.setitem(sys.modules, "PIL", fake_pil)

    class Dim:
        def __init__(self, width=None, height=None):
            self.width = width
            self.height = height

    class WS:
        def __init__(self):
            self.column_dimensions = {"C": Dim(width=8.43)}  # 近似 64px
            self.row_dimensions = {5: Dim(height=15.0)}       # 近似 20px

    ws = WS()
    w, h = calc_fit_size_openpyxl(ws, 5, 3, "img.png")
    assert w > 0 and h > 0
    # 宽度应与单元格宽度近似一致（充满宽度）
    assert abs(w - 64) < 3
