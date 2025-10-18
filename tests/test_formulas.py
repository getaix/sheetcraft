import sys
import types
import builtins

import pytest

from fexcel.formulas import evaluate_xlsx_formulas


def test_formula_evaluation_with_stub(monkeypatch, tmpfile):
    """通过桩替换 xlcalculator，验证公式评估流程。"""

    # 构造桩模块
    xlcalculator = types.ModuleType("xlcalculator")

    class Cell:
        def __init__(self, addr, formula=None):
            self.formula = formula

    class Sheet:
        def __init__(self):
            self.cells = {
                "A1": Cell("A1", formula="SUM(1,2)"),
                "B2": Cell("B2"),
            }

    class Model:
        def __init__(self):
            self.book = {"Sheet1": Sheet()}

    class ModelCompiler:
        def read_and_parse_archive(self, path):
            return Model()

    class Evaluator:
        def __init__(self, model):
            self.model = model

        def evaluate(self, addr):
            return 3.0

    xlcalculator.ModelCompiler = ModelCompiler
    xlcalculator.Evaluator = Evaluator
    monkeypatch.setitem(sys.modules, "xlcalculator", xlcalculator)

    res = evaluate_xlsx_formulas(tmpfile)
    assert ("Sheet1", "A1") in res
    assert res[("Sheet1", "A1")] == 3.0


def test_formula_evaluation_missing_dep_raises(monkeypatch, tmpfile):
    """当无法导入 xlcalculator 时应抛出 RuntimeError。"""
    orig_import = builtins.__import__

    def _fake_import(name, *args, **kwargs):
        if name == "xlcalculator":
            raise ImportError("missing xlcalculator")
        return orig_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", _fake_import)
    with pytest.raises(RuntimeError):
        evaluate_xlsx_formulas(tmpfile)
