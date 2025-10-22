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

# YEARFRAC 实现：与 Excel YEARFRAC 兼容（basis 0–4）
from datetime import date as _date, datetime as _datetime, timedelta as _timedelta
import calendar as _calendar
from typing import Union as _Union


def _to_date(x: _Union[_date, _datetime, str]) -> _date:
    if isinstance(x, _datetime):
        return x.date()
    if isinstance(x, _date):
        return x
    if isinstance(x, str):
        try:
            return _date.fromisoformat(x)
        except Exception as exc:
            raise TypeError("start_date/end_date 字符串仅支持 ISO 格式 YYYY-MM-DD") from exc
    raise TypeError("start_date/end_date 必须为 date、datetime 或 ISO 日期字符串")


def _is_leap_year(y: int) -> bool:
    return y % 4 == 0 and (y % 100 != 0 or y % 400 == 0)


def _is_end_of_month(d: _date) -> bool:
    return d.day == _calendar.monthrange(d.year, d.month)[1]


def _days360_us_nasd(d1: _date, d2: _date) -> int:
    y1, m1, dd1 = d1.year, d1.month, d1.day
    y2, m2, dd2 = d2.year, d2.month, d2.day

    # US (NASD) 调整规则，与 Excel YEARFRAC basis=0 对齐
    if dd1 == 31:
        dd1 = 30
    if m1 == 2 and _is_end_of_month(d1):
        dd1 = 30

    if dd2 == 31:
        if dd1 in (30, 31):
            dd2 = 30
    if m2 == 2 and _is_end_of_month(d2):
        # 若两者均为 2 月末，则按 30 调整
        if m1 == 2 and _is_end_of_month(d1):
            dd2 = 30

    return (360 * (y2 - y1)) + (30 * (m2 - m1)) + (dd2 - dd1)


def _days360_eu(d1: _date, d2: _date) -> int:
    y1, m1, dd1 = d1.year, d1.month, d1.day
    y2, m2, dd2 = d2.year, d2.month, d2.day
    dd1 = min(dd1, 30)
    dd2 = min(dd2, 30)
    return (360 * (y2 - y1)) + (30 * (m2 - m1)) + (dd2 - dd1)


def yearfrac(start_date: _Union[_date, _datetime, str], end_date: _Union[_date, _datetime, str], basis: int = 0) -> float:
    """计算两个日期之间的年分数（Excel YEARFRAC 兼容）。

    参数：
    - start_date, end_date：`date`、`datetime` 或 ISO 字符串 `YYYY-MM-DD`。
    - basis：日计数基准，0..4：
      0 US(NASD) 30/360；1 Actual/Actual；2 Actual/360；3 Actual/365；4 European 30/360。

    返回：
    - 浮点数，可能为负（当 end_date < start_date）。
    """
    if not isinstance(basis, int) or basis < 0 or basis > 4:
        raise ValueError("basis 必须为 0..4 之间的整数")

    d1 = _to_date(start_date)
    d2 = _to_date(end_date)

    if d1 == d2:
        return 0.0

    sign = 1.0
    if d2 < d1:
        d1, d2 = d2, d1
        sign = -1.0

    if basis == 0:
        return sign * (_days360_us_nasd(d1, d2) / 360.0)

    diff = (d2 - d1).days

    if basis == 2:
        return sign * (diff / 360.0)
    if basis == 3:
        return sign * (diff / 365.0)

    if basis == 4:
        return sign * (_days360_eu(d1, d2) / 360.0)

    # basis == 1: Actual/Actual（分年度片段，每段按当年天数归一化）
    total = 0.0
    cur = d1
    while cur < d2:
        # 当年结束（下一年的 1 月 1 日）或 d2，以较小者为段终
        next_year_start = _date(cur.year + 1, 1, 1)
        seg_end = d2 if d2 < next_year_start else next_year_start
        denom = 366.0 if _is_leap_year(cur.year) else 365.0
        seg_days = (seg_end - cur).days
        total += seg_days / denom
        cur = seg_end
    return sign * total
