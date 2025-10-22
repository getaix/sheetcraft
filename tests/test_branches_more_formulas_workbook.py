import os
import types
import builtins

import pytest

from sheetcraft.formulas import evaluate_xlsx_formulas
from sheetcraft.workbook import ExcelWorkbook, DataValidationSpec


def test_evaluate_xlsx_formulas_requires_xlcalculator(monkeypatch, tmp_path):
    # Force import error for xlcalculator to exercise RuntimeError branch
    orig_import = builtins.__import__

    def _fake_import(name, *args, **kwargs):
        if name == "xlcalculator":
            raise ImportError("missing xlcalculator")
        return orig_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", _fake_import)
    with pytest.raises(RuntimeError) as exc:
        evaluate_xlsx_formulas(str(tmp_path / "fake.xlsx"))
    assert "xlcalculator" in str(exc.value)


def test_openpyxl_data_validation_between_whole(tmp_path):
    out = tmp_path / "dv_openpyxl.xlsx"
    wb = ExcelWorkbook(output_path=str(out))
    ws = wb.sheet("DV")
    spec = DataValidationSpec(type="whole", operator="between", formula1="1", formula2="10")
    wb.add_data_validation(ws, "A1:A3", spec)
    # Verify validation added in openpyxl
    assert hasattr(ws, "data_validations")
    # openpyxl stores list in data_validations.dataValidation
    assert getattr(ws.data_validations, "dataValidation", [])
    wb.save()
    assert out.exists()


def test_openpyxl_write_only_row_streaming(tmp_path):
    out = tmp_path / "wo.xlsx"
    wb = ExcelWorkbook(output_path=str(out), write_only=True)
    ws = wb.sheet("WO")
    wb.write_row(ws, 1, ["A", 1, True], styles=[{"bold": True}, None, None])
    wb.write_row(ws, 2, ["B", 2, False])
    wb.save()
    assert out.exists()


def test_save_apply_format_fix_on_save(monkeypatch, tmp_path):
    # Stub fix_xlsx to ensure the post-save formatting branch executes
    from sheetcraft import format_fix as ff

    called = {"ok": False}

    def _stub_fix(path, *_args, **_kwargs):
        called["ok"] = True

    monkeypatch.setattr(ff, "fix_xlsx", _stub_fix)
    out = tmp_path / "fix.xlsx"
    wb = ExcelWorkbook(output_path=str(out), apply_format_fix_on_save=True)
    ws = wb.sheet("S")
    wb.write_cell(ws, 1, 1, "X")
    wb.save()
    assert out.exists()
    assert called["ok"] is True


def test_preview_temp_applies_format_fix(monkeypatch):
    # Stub fix_xlsx again
    from sheetcraft import format_fix as ff

    called = {"ok": False}

    def _stub_fix(path, *_args, **_kwargs):
        called["ok"] = True

    monkeypatch.setattr(ff, "fix_xlsx", _stub_fix)
    wb = ExcelWorkbook(apply_format_fix_on_save=True)
    wb.write_cell("T", 1, 1, "preview")
    with wb.preview_temp() as p:
        assert os.path.exists(p)
    assert called["ok"] is True
    assert not os.path.exists(p)