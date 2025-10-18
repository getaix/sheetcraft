import os
import sys
import tempfile
import types
from pathlib import Path

import pytest

# 确保项目根目录加入导入路径，避免在某些运行环境下出现 ModuleNotFoundError
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


@pytest.fixture
def tmpfile():
    """临时文件路径生成器。"""
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        yield path
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


@pytest.fixture
def tmppath_dir(tmp_path):
    """临时目录路径。"""
    return tmp_path


def stub_xlsxwriter(monkeypatch):
    """创建一个 xlsxwriter 的最小桩以覆盖 fast 引擎路径。"""
    xlsxwriter = types.ModuleType("xlsxwriter")

    class FakeFormat(dict):
        pass

    class FakeWorksheet:
        def __init__(self, name="Sheet1"):
            self.name = name
            self._data = {}

        def write(self, r, c, v, fmt=None):
            self._data[(r, c)] = v

        def write_formula(self, r, c, v, fmt=None):
            self._data[(r, c)] = v

        def set_row(self, r, height=None):
            pass

        def set_column(self, c1, c2, width):
            pass

        def merge_range(self, r1, c1, r2, c2, value, fmt=None):
            # no-op for testing
            pass

        def insert_image(self, r, c, path, opts=None):
            # no-op
            pass

        def data_validation(self, r1, c1, r2, c2, options):
            # record options for assertion if needed
            self._dv = (r1, c1, r2, c2, options)

    class FakeWorkbook:
        def __init__(self, path):
            self.path = path
            self._formats = []

        def add_worksheet(self, name):
            return FakeWorksheet(name)

        def add_format(self, d):
            f = FakeFormat(d)
            self._formats.append(f)
            return f

        def close(self):
            # 在桩环境中生成空文件以便测试断言
            try:
                with open(self.path, "wb") as f:
                    f.write(b"")
            except Exception:
                pass

    xlsxwriter.Workbook = FakeWorkbook
    monkeypatch.setitem(sys.modules, "xlsxwriter", xlsxwriter)
    return xlsxwriter


def stub_xlwt(monkeypatch):
    """创建一个 xlwt 的最小桩以覆盖 .xls 路径。"""
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
            # 生成空文件以便测试断言
            try:
                with open(path, "wb") as f:
                    f.write(b"")
            except Exception:
                pass

    xlwt.Workbook = Workbook
    xlwt.XFStyle = XFStyle
    xlwt.Font = Font
    xlwt.Alignment = Alignment
    xlwt.Pattern = Pattern
    xlwt.Borders = Borders
    xlwt.Formula = Formula
    monkeypatch.setitem(sys.modules, "xlwt", xlwt)
    return xlwt
