import os
import types
import builtins

import pytest

from fexcel.workbook import ExcelWorkbook, DataValidationSpec
from fexcel.utils import get_column_letter_safe, convert_image_to_bmp_temp


def _stub_xlsxwriter(monkeypatch):
    xlsxwriter = types.ModuleType("xlsxwriter")

    class FakeFormat(dict):
        pass

    class FakeWorksheet:
        def __init__(self, name="Sheet1"):
            self.name = name
            self._data = {}
            self._dv = None

        def write(self, r, c, v, fmt=None):
            self._data[(r, c)] = v

        def write_formula(self, r, c, v, fmt=None):
            self._data[(r, c)] = v

        def set_row(self, r, height=None):
            pass

        def set_column(self, c1, c2, width):
            pass

        def merge_range(self, r1, c1, r2, c2, value, fmt=None):
            pass

        def insert_image(self, r, c, path, opts=None):
            pass

        def data_validation(self, r1, c1, r2, c2, options):
            self._dv = (r1, c1, r2, c2, options)

    class FakeWorkbook:
        def __init__(self, target, options=None):
            self.target = target
            self.options = options or {}

        def add_worksheet(self, name):
            return FakeWorksheet(name)

        def add_format(self, d):
            return FakeFormat(d)

        def close(self):
            data = b"PK"  # zip header prefix to mimic xlsx content
            try:
                if hasattr(self.target, "write"):
                    self.target.write(data)
                    # rewind for readers
                    if hasattr(self.target, "seek"):
                        self.target.seek(0)
                else:
                    with open(self.target, "wb") as f:
                        f.write(data)
            except Exception:
                pass

    xlsxwriter.Workbook = FakeWorkbook
    monkeypatch.setitem(__import__("sys").modules, "xlsxwriter", xlsxwriter)
    return xlsxwriter


def _stub_xlwt(monkeypatch):
    xlwt = types.ModuleType("xlwt")

    class Font:
        UNDERLINE_SINGLE = 1

        def __init__(self):
            self.bold = False
            self.italic = False
            self.underline = 0
            self.height = 200
            self.name = "Arial"

    class Alignment:
        HORZ_LEFT = 1
        HORZ_CENTER = 2
        HORZ_RIGHT = 3
        HORZ_GENERAL = 4
        VERT_TOP = 1
        VERT_CENTER = 2
        VERT_BOTTOM = 3

        def __init__(self):
            self.horz = self.HORZ_GENERAL
            self.vert = self.VERT_CENTER
            self.wrap = 0

    class Pattern:
        SOLID_PATTERN = 1

        def __init__(self):
            self.pattern = 0

    class Borders:
        THIN = 1

        def __init__(self):
            self.left = self.right = self.top = self.bottom = 0

    class XFStyle:
        def __init__(self):
            self.font = None
            self.alignment = None
            self.pattern = None
            self.borders = None

    class Formula:
        def __init__(self, f):
            self.f = f

    class Row:
        def __init__(self):
            self.height = 0

    class Col:
        def __init__(self):
            self.width = 0

    class Worksheet:
        def __init__(self, name):
            self.name = name
            self._data = {}
            self._rows = {}
            self._cols = {}

        def write(self, r, c, v, fmt=None):
            self._data[(r, c)] = v

        def row(self, r):
            self._rows.setdefault(r, Row())
            return self._rows[r]

        def col(self, c):
            self._cols.setdefault(c, Col())
            return self._cols[c]

        def merge(self, r1, r2, c1, c2):
            pass

        def insert_bitmap(self, path, r, c):
            pass

    class Workbook:
        def __init__(self, encoding="utf-8"):
            self._sheets = []

        def add_sheet(self, name):
            ws = Worksheet(name)
            self._sheets.append(ws)
            return ws

        def save(self, path):
            try:
                data = b"XLS"  # minimal bytes to mimic xls content
                # Support saving to file-like buffers for in-memory tests
                if hasattr(path, "write"):
                    path.write(data)
                    if hasattr(path, "seek"):
                        path.seek(0)
                else:
                    with open(path, "wb") as f:
                        f.write(data)
            except Exception:
                pass

    xlwt.Workbook = Workbook
    xlwt.XFStyle = XFStyle
    xlwt.Font = Font
    xlwt.Alignment = Alignment
    xlwt.Pattern = Pattern
    xlwt.Borders = Borders
    xlwt.Formula = Formula
    monkeypatch.setitem(__import__("sys").modules, "xlwt", xlwt)
    return xlwt


def test_openpyxl_write_only_row_styles(tmp_path):
    out = tmp_path / "wo.xlsx"
    wb = ExcelWorkbook(output_path=str(out), write_only=True)
    ws = wb.sheet()
    wb.write_row(
        ws,
        1,
        ["A", "B", "C"],
        styles=[{"bold": True}, {"italic": True}, {"underline": True}],
    )
    wb.save()
    assert out.exists()


def test_xlsxwriter_engine_full_flow(monkeypatch, tmp_path):
    _stub_xlsxwriter(monkeypatch)
    out = tmp_path / "fast.xlsx"
    wb = ExcelWorkbook(output_path=str(out), fast=True)
    ws = wb.add_sheet("S")
    # write and formula
    wb.write_cell(ws, 1, 1, "Hello")
    wb.write_cell(ws, 2, 1, "=SUM(A1:A1)")
    # layout
    wb.set_row_height(ws, 1, 20)
    wb.set_column_width(ws, 1, 12)
    # merge
    wb.merge_cells(ws, 3, 1, 3, 2, value="Merged")
    # image insertion
    wb.insert_image(ws, 4, 1, str(tmp_path / "fake.png"))
    # data validation (list)
    spec = DataValidationSpec(type="list", formula1="A1:A10")
    wb.add_data_validation(ws, "A5:A10", spec)
    # set formula via helper
    wb.set_formula(ws, 5, 1, "SUM(A1:A2)")
    wb.save()
    assert out.exists()
    # validate that data_validation recorded options
    assert hasattr(ws, "_dv") and ws._dv[4]["validate"] == "list"


def test_xlsxwriter_in_memory_get_bytes(monkeypatch):
    _stub_xlsxwriter(monkeypatch)
    wb = ExcelWorkbook(fast=True)
    ws = wb.add_sheet("S")
    wb.write_cell(ws, 1, 1, "Hi")
    data = wb.get_bytes()
    assert isinstance(data, (bytes, bytearray)) and len(data) > 0


def test_xlwt_engine_basic_ops(monkeypatch, tmp_path):
    _stub_xlwt(monkeypatch)
    out = tmp_path / "book.xls"
    wb = ExcelWorkbook(output_path=str(out), file_format="xls")
    ws = wb.add_sheet("S1")
    wb.write_cell(ws, 1, 1, "X")
    wb.set_row_height(ws, 1, 18)
    wb.set_column_width(ws, 1, 20)
    wb.merge_cells(ws, 2, 1, 2, 2, value="M")
    # image path can be anything; conversion errors are gracefully handled in implementation
    wb.insert_image(ws, 3, 1, str(tmp_path / "fake.jpg"))
    wb.set_formula(ws, 4, 1, "A1+A2")
    wb.save()
    assert out.exists()


def test_parse_range_and_resolve_sheet():
    wb = ExcelWorkbook()
    # _parse_range single and range
    r1, c1, r2, c2 = wb._parse_range("B2")
    assert (r1, c1, r2, c2) == (2, 2, 2, 2)
    r1, c1, r2, c2 = wb._parse_range("A1:D10")
    assert (r1, c1, r2, c2) == (1, 1, 10, 4)
    # resolve by name auto-creates
    wb.write_cell("V", 1, 1, "val")
    assert "V" in wb._sheets


def test_write_table_and_export_dicts(tmp_path):
    out = tmp_path / "dicts.xlsx"
    wb = ExcelWorkbook(output_path=str(out))
    data = [
        {"a": 1, "b": 2},
        {"a": 3, "b": 4},
    ]
    header_map = {"a": "A列", "b": "B列"}
    wb.export_dicts(
        "Report",
        1,
        data,
        header_map=header_map,
        order=["b", "a"],
        header_style={"bold": True},
        row_style={"italic": True},
    )
    # also cover empty data early return
    wb.export_dicts("Report", 10, [])
    wb.save()
    assert out.exists()


def test_validation_xlsxwriter_mapping_variants(monkeypatch, tmp_path):
    _stub_xlsxwriter(monkeypatch)
    out = tmp_path / "dv.xlsx"
    wb = ExcelWorkbook(output_path=str(out), fast=True)
    ws = wb.add_sheet("DV")
    # between mapping with numeric type
    spec = DataValidationSpec(
        type="whole", operator="between", formula1="1", formula2="10"
    )
    wb.add_data_validation(ws, "A1:A5", spec)
    dv_opts = ws._dv[4]
    assert dv_opts["validate"] == "whole"
    assert dv_opts["minimum"] == "1" and dv_opts["maximum"] == "10"
    # notBetween variant (camel)
    spec2 = DataValidationSpec(
        type="whole", operator="notBetween", formula1="5", formula2="15"
    )
    wb.add_data_validation(ws, "B1:B5", spec2)
    dv_opts2 = ws._dv[4]
    assert dv_opts2["maximum"] == "15"
    # plain comparison
    spec3 = DataValidationSpec(type="decimal", operator=">=", formula1="3")
    wb.add_data_validation(ws, "C1:C5", spec3)
    dv_opts3 = ws._dv[4]
    assert dv_opts3["criteria"] == ">=" and dv_opts3["value"] == "3"
    wb.save()
    assert out.exists()


def test_utils_get_column_letter_safe_fallback(monkeypatch):
    # Make openpyxl.utils import succeed but missing attribute to trigger except branch
    fake_utils = types.ModuleType("openpyxl.utils")
    monkeypatch.setitem(__import__("sys").modules, "openpyxl.utils", fake_utils)
    letter = get_column_letter_safe(3)
    assert letter == "C"

    # Trigger raise branch when index > 26
    with pytest.raises(ImportError):
        get_column_letter_safe(27)


def test_utils_convert_image_to_bmp_temp_with_stub(monkeypatch, tmp_path):
    # Stub Pillow's Image module
    pil = types.ModuleType("PIL")
    img_mod = types.ModuleType("Image")

    class _FakeImg:
        def convert(self, mode):
            return self

        def save(self, path, format=None):
            with open(path, "wb") as f:
                f.write(b"BM")

    def _open(path):
        return _FakeImg()

    img_mod.open = _open
    pil.Image = img_mod
    monkeypatch.setitem(__import__("sys").modules, "PIL", pil)
    monkeypatch.setitem(__import__("sys").modules, "PIL.Image", img_mod)

    # Use any existing file path; create a dummy png
    png = tmp_path / "x.png"
    with open(png, "wb") as f:
        f.write(b"\x89PNG")
    bmp_path = convert_image_to_bmp_temp(str(png))
    assert os.path.exists(bmp_path)
    os.remove(bmp_path)


def test_utils_convert_image_to_bmp_temp_raises_without_pillow(monkeypatch, tmp_path):
    # Force ImportError for PIL
    orig_import = builtins.__import__

    def _fake_import(name, *args, **kwargs):
        if name.startswith("PIL"):
            raise ImportError("No PIL")
        return orig_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", _fake_import)

    png = tmp_path / "y.png"
    with open(png, "wb") as f:
        f.write(b"\x89PNG")
    with pytest.raises(RuntimeError):
        convert_image_to_bmp_temp(str(png))


def test_openpyxl_set_column_width_and_infer_unknown_extension(tmp_path):
    # Unknown extension falls back to openpyxl
    out = tmp_path / "file.unknown"
    wb = ExcelWorkbook(output_path=str(out))
    ws = wb.sheet()
    wb.set_column_width(ws, 1, 22)
    wb.save()
    assert out.exists()


def test_preview_temp_creates_and_deletes(tmp_path):
    wb = ExcelWorkbook()
    wb.write_cell("T", 1, 1, "temp")
    with wb.preview_temp() as p:
        assert os.path.exists(p)
    assert not os.path.exists(p)


def test_xlsxwriter_get_sheet_default_creation(monkeypatch, tmp_path):
    _stub_xlsxwriter(monkeypatch)
    wb = ExcelWorkbook(output_path=str(tmp_path / "def.xlsx"), fast=True)
    # no sheets exist yet; get_sheet(None) should create default
    ws = wb.get_sheet()
    assert ws is not None


def test_xlwt_import_failure_raises(monkeypatch, tmp_path):
    # Force xlwt import failure and verify helpful RuntimeError
    orig_import = builtins.__import__

    def _fake_import(name, *args, **kwargs):
        if name == "xlwt":
            raise ImportError("missing xlwt")
        return orig_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", _fake_import)
    with pytest.raises(RuntimeError):
        ExcelWorkbook(output_path=str(tmp_path / "bad.xls"), file_format="xls")


def test_validation_xlsxwriter_custom_cleans_none(monkeypatch, tmp_path):
    _stub_xlsxwriter(monkeypatch)
    out = tmp_path / "dv2.xlsx"
    wb = ExcelWorkbook(output_path=str(out), fast=True)
    ws = wb.add_sheet("DV2")
    spec = DataValidationSpec(type="custom", formula1="A1>0")
    wb.add_data_validation(ws, "A1:A1", spec)
    opts = ws._dv[4]
    # ensure no None values present in options
    assert all(v is not None for v in opts.values())
    wb.save()
    assert out.exists()


def test_openpyxl_get_bytes_and_merge_style():
    wb = ExcelWorkbook()
    ws = wb.sheet()
    wb.merge_cells(
        ws,
        1,
        1,
        1,
        2,
        value="V",
        style={
            "font": {"bold": True, "italic": True},
            "align": {"horizontal": "center"},
        },
    )
    data = wb.get_bytes()
    assert isinstance(data, (bytes, bytearray)) and len(data) > 0


def test_openpyxl_get_sheet_active_default():
    wb = ExcelWorkbook()
    ws = wb.get_sheet()
    # active sheet should exist and be returned
    assert ws is not None and getattr(ws, "title", "")


def test_xlsxwriter_preview_temp_in_memory(monkeypatch):
    _stub_xlsxwriter(monkeypatch)
    wb = ExcelWorkbook(fast=True)
    wb.write_cell("P", 1, 1, "pv")
    with wb.preview_temp() as p:
        assert os.path.exists(p) and p.endswith(".xlsx")
    assert not os.path.exists(p)


def test_xlsxwriter_get_bytes_path_mode(monkeypatch, tmp_path):
    _stub_xlsxwriter(monkeypatch)
    out = tmp_path / "pm.xlsx"
    wb = ExcelWorkbook(output_path=str(out), fast=True)
    ws = wb.add_sheet("S")
    wb.write_cell(ws, 1, 1, "x")
    data = wb.get_bytes()
    assert isinstance(data, (bytes, bytearray)) and len(data) > 0
    assert out.exists()


def test_xlwt_in_memory_get_bytes(monkeypatch):
    _stub_xlwt(monkeypatch)
    wb = ExcelWorkbook(file_format="xls")
    ws = wb.add_sheet("S")
    wb.write_cell(ws, 1, 1, "old")
    b = wb.get_bytes()
    assert isinstance(b, (bytes, bytearray)) and len(b) > 0
