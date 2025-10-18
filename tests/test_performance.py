import pytest

from sheetcraft import ExcelWorkbook


@pytest.mark.performance
def test_perf_openpyxl_write_5000_rows(benchmark, tmp_path):
    """基准测试：openpyxl 写入 5000 行。"""
    out = tmp_path / "perf_openpyxl.xlsx"
    rows = [[i, i * 2, i * 3.14, f"text-{i}"] for i in range(1, 5001)]

    def run():
        wb = ExcelWorkbook(output_path=str(out), fast=False)
        ws = wb.sheet("Perf")
        wb.write_table(ws, 1, ["A", "B", "C", "D"], rows)
        wb.save()

    benchmark(run)
    assert out.exists()


@pytest.mark.performance
def test_perf_xlsxwriter_fast_5000_rows(benchmark, monkeypatch, tmp_path):
    """基准测试：xlsxwriter 快速写入 5000 行（通过桩）。"""
    from tests.conftest import stub_xlsxwriter

    stub_xlsxwriter(monkeypatch)
    out = tmp_path / "perf_fast.xlsx"
    rows = [[i, i * 2, i * 3.14, f"text-{i}"] for i in range(1, 5001)]

    def run():
        wb = ExcelWorkbook(output_path=str(out), fast=True)
        ws = wb.add_sheet("Perf")
        wb.write_table(ws, 1, ["A", "B", "C", "D"], rows)
        wb.save()

    benchmark(run)
    assert out.exists()


@pytest.mark.performance
def test_stress_openpyxl_write_10000_rows(tmp_path):
    """压力测试：openpyxl 写入 10000 行，记录耗时并验证输出。"""
    import time

    out = tmp_path / "stress_openpyxl.xlsx"
    rows = [[i, i * 2, i * 3.14, f"text-{i}"] for i in range(1, 10001)]
    start = time.time()
    wb = ExcelWorkbook(output_path=str(out), fast=False)
    ws = wb.sheet("Stress")
    wb.write_table(ws, 1, ["A", "B", "C", "D"], rows)
    wb.save()
    elapsed = time.time() - start
    assert out.exists()
    # 适度阈值，避免测试环境波动；记录而非严格限制
    assert elapsed < 60
