import os
import io
import types
import pytest

from sheetcraft.workbook import ExcelWorkbook, DataValidationSpec


def test_infer_format_xls_without_file_format(monkeypatch, tmp_path):
    from tests.conftest import stub_xlwt

    stub_xlwt(monkeypatch)
    out = tmp_path / "infer_format.xls"
    wb = ExcelWorkbook(output_path=str(out))
    ws = wb.add_sheet("S")
    wb.write_cell(ws, 1, 1, "XLS")
    wb.save()
    assert out.exists()


def test_get_sheet_first_returns_existing_xlsxwriter(monkeypatch, tmp_path):
    from tests.conftest import stub_xlsxwriter

    stub_xlsxwriter(monkeypatch)
    out = tmp_path / "first.xlsx"
    wb = ExcelWorkbook(output_path=str(out), fast=True)
    ws1 = wb.add_sheet("A")
    ws2 = wb.add_sheet("B")
    first = wb.get_sheet()
    assert first is ws1
    wb.save()
    assert out.exists()


def test_get_sheet_existing_name_returns(monkeypatch, tmp_path):
    from tests.conftest import stub_xlsxwriter

    stub_xlsxwriter(monkeypatch)
    out = tmp_path / "byname.xlsx"
    wb = ExcelWorkbook(output_path=str(out), fast=True)
    ws = wb.add_sheet("S")
    got = wb.get_sheet("S")
    assert got is ws
    wb.save()
    assert out.exists()


def test_openpyxl_set_formula_with_style(tmp_path):
    out = tmp_path / "formula_style.xlsx"
    wb = ExcelWorkbook(output_path=str(out))
    ws = wb.add_sheet("Calc")
    wb.set_formula(ws, 1, 1, "SUM(1,2)", style={"font": {"bold": True}})
    wb.save()
    assert out.exists()


def test_xlwt_data_validation_noop(monkeypatch, tmp_path):
    from tests.conftest import stub_xlwt

    stub_xlwt(monkeypatch)
    out = tmp_path / "dv_noop.xls"
    wb = ExcelWorkbook(output_path=str(out))
    ws = wb.add_sheet("S")
    spec = DataValidationSpec(type="list", formula1="A,B", allow_blank=True)
    # xlwt branch is a no-op; should not raise
    wb.add_data_validation(ws, "A1:A1", spec)
    wb.save()
    assert out.exists()


def test_save_apply_format_fix_failure(monkeypatch, tmp_path):
    # Patch fix_xlsx to raise so the except path is covered
    import sheetcraft.format_fix as ff

    def boom(*args, **kwargs):
        raise RuntimeError("boom")

    monkeypatch.setattr(ff, "fix_xlsx", boom, raising=True)

    out = tmp_path / "save_fix_fail.xlsx"
    wb = ExcelWorkbook(output_path=str(out), apply_format_fix_on_save=True)
    ws = wb.add_sheet("S")
    wb.write_cell(ws, 1, 1, "x")
    wb.save()
    assert out.exists()


def test_preview_temp_format_fix_failure_and_remove_raises(monkeypatch, tmp_path):
    # Patch fix_xlsx to raise to cover preview_temp's except branch
    import sheetcraft.format_fix as ff

    def boom(*args, **kwargs):
        raise RuntimeError("boom")

    monkeypatch.setattr(ff, "fix_xlsx", boom, raising=True)

    wb = ExcelWorkbook(apply_format_fix_on_save=True)
    ws = wb.add_sheet("S")
    wb.write_cell(ws, 1, 1, "x")

    with wb.preview_temp() as p:
        assert os.path.exists(p)
        # file should exist during context
        with open(p, "rb") as f:
            assert f.read(4) in {b"PK\x03\x04", b"<?xm"}


def test_export_dicts_order_none(tmp_path):
    out = tmp_path / "export_order_none.xlsx"
    wb = ExcelWorkbook(output_path=str(out))
    ws = wb.add_sheet("Data")
    data = [
        {"a": 1, "b": 2},
        {"a": 3, "b": 4},
    ]
    # order=None triggers default order branch
    wb.export_dicts(ws, 1, data, header_map=None, order=None)
    wb.save()
    assert out.exists()