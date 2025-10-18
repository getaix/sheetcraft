import os
import sys
import types


from sheetcraft.utils import (
    ensure_parent_dir,
    get_column_letter_safe,
    convert_image_to_bmp_temp,
)
from sheetcraft.styles import (
    apply_openpyxl_style,
    xlsxwriter_format_from_dict,
    xlwt_style_from_dict,
)


def test_ensure_parent_dir(tmp_path):
    """验证 ensure_parent_dir 创建父目录。"""
    target = tmp_path / "a" / "b" / "c.xlsx"
    ensure_parent_dir(str(target))
    assert (tmp_path / "a" / "b").exists()


def test_get_column_letter_safe_fallback(monkeypatch):
    """通过桩替换 openpyxl.utils 缺失，触发 fallback。"""
    fake_mod = types.ModuleType("openpyxl.utils")
    # 不提供 get_column_letter 以触发 ImportError
    monkeypatch.setitem(sys.modules, "openpyxl.utils", fake_mod)
    assert get_column_letter_safe(1) == "A"


def test_xlsxwriter_format_from_dict_basic():
    """验证 xlsxwriter 格式转换。"""

    class WB:
        def add_format(self, d):
            return d

    style = {
        "font": {"bold": True, "size": 12},
        "align": {"horizontal": "center"},
        "fill": {"color": "#ffeecc"},
        "number_format": "#,##0.00",
    }
    fmt = xlsxwriter_format_from_dict(WB(), style)
    assert fmt["bold"] and fmt["font_size"] == 12
    assert fmt["align"] == "center"
    assert fmt["bg_color"] == "#ffeecc"
    assert fmt["num_format"] == "#,##0.00"


def test_xlwt_style_from_dict_with_stub(monkeypatch):
    """通过 xlwt 桩验证 XFStyle 构造。"""
    from tests.conftest import stub_xlwt

    stub_xlwt(monkeypatch)
    style = {
        "font": {"bold": True, "size": 12},
        "align": {"horizontal": "center", "vertical": "bottom", "wrap": True},
        "border": {"left": True, "right": True},
    }
    xf = xlwt_style_from_dict(style)
    assert xf.font is not None and xf.alignment is not None


def test_apply_openpyxl_style_with_stub(monkeypatch):
    """通过 openpyxl.styles 桩验证样式应用。"""
    styles_mod = types.ModuleType("openpyxl.styles")

    class Font:
        def __init__(self, **kwargs):
            self.kwargs = kwargs

    class Alignment:
        def __init__(self, **kwargs):
            self.kwargs = kwargs

    class PatternFill:
        def __init__(self, **kwargs):
            self.kwargs = kwargs

    class Side:
        def __init__(self, **kwargs):
            self.kwargs = kwargs

    class Border:
        def __init__(self, **kwargs):
            self.kwargs = kwargs

    styles_mod.Font = Font
    styles_mod.Alignment = Alignment
    styles_mod.PatternFill = PatternFill
    styles_mod.Side = Side
    styles_mod.Border = Border
    monkeypatch.setitem(sys.modules, "openpyxl.styles", styles_mod)

    class Cell:
        font = alignment = fill = border = number_format = None

    c = Cell()
    style = {
        "font": {"bold": True, "color": "#000000"},
        "align": {"horizontal": "center"},
        "fill": {"color": "#ffeecc"},
        "border": {"left": True, "right": True},
        "number_format": "#,##0.0",
    }
    apply_openpyxl_style(c, style)
    assert c.font and c.alignment and c.fill and c.border and c.number_format


def test_convert_image_to_bmp_temp(monkeypatch, tmp_path):
    """通过 Pillow 桩验证 BMP 转换逻辑。"""
    fake_img_mod = types.ModuleType("PIL.Image")

    class FakeImageObj:
        def convert(self, mode):
            return self

        def save(self, path, format=None):
            # 创建一个空文件以模拟保存
            with open(path, "wb") as f:
                f.write(b"BM")

    class FakeImage:
        @staticmethod
        def open(path):
            return FakeImageObj()

    fake_img_mod.Image = FakeImage
    # PIL 包模块结构：PIL.Image 中定义 Image 类，但我们以模块方式注入
    monkeypatch.setitem(sys.modules, "PIL", types.ModuleType("PIL"))
    monkeypatch.setitem(sys.modules, "PIL.Image", FakeImage)

    src = tmp_path / "src.png"
    src.write_bytes(b"\x89PNG")
    out = convert_image_to_bmp_temp(str(src))
    assert out.endswith(".bmp") and os.path.exists(out)
