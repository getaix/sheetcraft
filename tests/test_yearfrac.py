from datetime import date
import pytest

from sheetcraft.formulas import yearfrac


def test_yearfrac_basis1_non_leap_short():
    assert yearfrac(date(2015, 4, 20), date(2015, 5, 1), 1) == pytest.approx(11 / 365)


def test_yearfrac_basis1_leap_short():
    assert yearfrac(date(2016, 4, 20), date(2016, 5, 1), 1) == pytest.approx(11 / 366)


def test_yearfrac_basis1_cross_year_segments():
    # 2016 是闰年，跨年分段：2016-02-20 到 2017-01-01 -> 316/366
    assert yearfrac(date(2016, 2, 20), date(2017, 1, 1), 1) == pytest.approx(316 / 366)


def test_yearfrac_basis1_multi_year_segments_with_string_input():
    # 2015-12-31 到 2018-01-02：分段 2015(1/365) + 2016(366/366) + 2017(365/365) + 2018(1/365)
    # 验证字符串输入被正确解析
    v = yearfrac("2015-12-31", "2018-01-02", 1)
    expect = (1/365) + 1.0 + 1.0 + (1/365)
    assert v == pytest.approx(expect)


def test_yearfrac_basis0_us_30_360_half_year():
    # 与文档示例一致：2008-01-01 到 2008-07-01 -> 0.5（US 30/360）
    assert yearfrac(date(2008, 1, 1), date(2008, 7, 1), 0) == pytest.approx(0.5)


def test_yearfrac_basis4_eu_30_360_half_year():
    assert yearfrac(date(2008, 1, 1), date(2008, 7, 1), 4) == pytest.approx(0.5)


def test_yearfrac_basis2_actual_360():
    assert yearfrac(date(2024, 1, 1), date(2024, 1, 12), 2) == pytest.approx(11 / 360)


def test_yearfrac_basis3_actual_365():
    assert yearfrac(date(2024, 1, 1), date(2024, 1, 12), 3) == pytest.approx(11 / 365)


def test_yearfrac_negative_when_end_before_start():
    assert yearfrac(date(2024, 1, 12), date(2024, 1, 1), 3) == pytest.approx(-11 / 365)


def test_yearfrac_invalid_basis_raises():
    with pytest.raises(ValueError):
        yearfrac(date(2024, 1, 1), date(2024, 1, 12), 5)