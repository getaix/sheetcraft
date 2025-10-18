from __future__ import annotations

from typing import Dict, Tuple


def evaluate_xlsx_formulas(path: str) -> Dict[Tuple[str, str], float]:
    """基于 xlcalculator（可选安装）评估 `.xlsx` 工作簿中的公式。

    返回形如 `{(sheet_name, cell_ref): value}` 的映射。
    若未安装 xlcalculator，则抛出带安装提示的 RuntimeError。
    """
    try:
        from xlcalculator import ModelCompiler, Evaluator  # type: ignore
    except Exception as exc:
        raise RuntimeError(
            "公式评估依赖 xlcalculator。请安装：pip install 'sheetcraft[formula]'"
        ) from exc

    mc = ModelCompiler()
    model = mc.read_and_parse_archive(path)
    ev = Evaluator(model)

    results: Dict[Tuple[str, str], float] = {}
    # 收集带公式的单元格并尝试评估
    for sheet_name, sheet in model.book.items():
        for addr, cell in sheet.cells.items():
            if getattr(cell, "formula", None):
                try:
                    val = ev.evaluate(f"{sheet_name}!{addr}")
                    results[(sheet_name, addr)] = val
                except Exception:
                    # 跳过评估失败的单元格
                    continue
    return results
