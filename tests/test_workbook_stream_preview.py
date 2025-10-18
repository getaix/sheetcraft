import os

from fexcel import ExcelWorkbook


def test_openpyxl_write_only_mode_streams(tmpfile):
    """覆盖 openpyxl 写入的 write_only 流式路径（使用 WriteOnlyCell）。"""
    wb = ExcelWorkbook(output_path=tmpfile, write_only=True)
    ws = wb.sheet("S")
    wb.write_row(ws, 1, ["A", 1, 2.5])
    wb.write_row(ws, 2, ["B", 2, 3.5])
    wb.save()
    assert os.path.exists(tmpfile)


def test_preview_temp_with_format_fix_enabled():
    """覆盖 preview_temp 的字节写入与可选格式修复分支。"""
    wb = ExcelWorkbook(apply_format_fix_on_save=True)
    ws = wb.sheet("P")
    wb.write_row(ws, 1, ["X", 1])
    with wb.preview_temp() as path:
        assert os.path.exists(path)
        # 读取部分字节，确认文件可用
        with open(path, "rb") as f:
            data = f.read(32)
            assert isinstance(data, (bytes, bytearray))
            assert len(data) > 0


def test_get_bytes_openpyxl_returns_data():
    """覆盖 get_bytes 的 openpyxl 分支。"""
    wb = ExcelWorkbook()
    ws = wb.sheet("G")
    wb.write_row(ws, 1, ["Y", 2])
    blob = wb.get_bytes()
    assert isinstance(blob, (bytes, bytearray))
    assert len(blob) > 0


def test_workbook_save_with_format_fix_on_save(tmpfile):
    """覆盖 save 后的可选格式修复分支（openpyxl）。"""
    wb = ExcelWorkbook(output_path=tmpfile, apply_format_fix_on_save=True)
    ws = wb.sheet("F")
    wb.write_row(ws, 1, ["Z", 3])
    wb.save()
    assert os.path.exists(tmpfile)