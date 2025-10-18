import os
from fexcel import ExcelWorkbook
from fexcel.workbook import DataValidationSpec


def test_openpyxl_write_and_style(tmpfile):
    """验证 openpyxl 写入与样式应用。"""
    wb = ExcelWorkbook(output_path=tmpfile, fast=False)
    ws = wb.sheet("Sheet1")
    style = {
        "font": {"bold": True, "size": 11},
        "align": {"horizontal": "center"},
        "fill": {"color": "#FFEECC"},
        "border": {"left": True, "right": True, "top": True, "bottom": True},
        "number_format": "#,##0.00",
    }
    wb.write_cell(ws, 1, 1, 123.456, style)
    wb.set_row_height(ws, 1, 22)
    wb.set_column_width(ws, 1, 18)
    wb.merge_cells(ws, 2, 1, 2, 3, value="Merged")
    wb.set_formula(ws, 3, 1, "SUM(1,2,3)")
    wb.save()

    assert os.path.exists(tmpfile)


def test_semantic_write_table_and_export_dicts(tmpfile):
    """验证语义化 API：write_table 与 export_dicts。"""
    wb = ExcelWorkbook(output_path=tmpfile)
    ws = wb.sheet("Report")
    headers = ["Name", "Qty", "Price"]
    rows = [["A", 1, 1.11], ["B", 2, 2.22]]
    wb.write_table(ws, 1, headers, rows)

    data = [
        {"name": "C", "qty": 3, "price": 3.33},
        {"name": "D", "qty": 4, "price": 4.44},
    ]
    wb.export_dicts(
        ws,
        4,
        data,
        header_map={"name": "Name", "qty": "Qty", "price": "Price"},
        order=["name", "qty", "price"],
    )
    wb.save()
    assert os.path.exists(tmpfile)


def test_openpyxl_validation(tmpfile):
    """验证 openpyxl 数据有效性。"""
    wb = ExcelWorkbook(output_path=tmpfile)
    ws = wb.sheet("V")
    spec = DataValidationSpec(type="list", formula1='"A,B,C"')
    wb.add_data_validation(ws, "A1:A10", spec)
    wb.save()
    assert os.path.exists(tmpfile)


def test_fast_xlsxwriter_paths(monkeypatch, tmpfile):
    """通过桩模块覆盖 xlsxwriter，验证快速写入分支。"""
    from tests.conftest import stub_xlsxwriter

    stub_xlsxwriter(monkeypatch)

    wb = ExcelWorkbook(output_path=tmpfile, fast=True)
    ws = wb.add_sheet("Fast")
    wb.write_row(ws, 1, ["X", 1, "=SUM(1,2)"])
    wb.add_data_validation(
        ws, "A1:A10", DataValidationSpec(type="whole", operator=">=", formula1="0")
    )
    wb.save()
    assert os.path.exists(tmpfile)


def test_xlwt_legacy_paths(monkeypatch, tmp_path):
    """通过桩模块覆盖 xlwt，验证 .xls 分支。"""
    from tests.conftest import stub_xlwt

    stub_xlwt(monkeypatch)

    out = tmp_path / "legacy.xls"
    wb = ExcelWorkbook(output_path=str(out), file_format="xls")
    ws = wb.add_sheet("L")
    wb.write_row(ws, 1, ["A", 1, 2])
    wb.set_row_height(ws, 1, 22)
    wb.set_column_width(ws, 1, 18)
    wb.merge_cells(ws, 2, 1, 2, 2, value="M")
    wb.set_formula(ws, 3, 1, "SUM(1,2)")
    # 图片插入走桩，无实际转换
    wb.insert_image(ws, 1, 3, "fake.jpg")
    wb.save(str(out))
    assert out.exists()


def test_write_rows_batch(tmpfile):
    """覆盖 write_rows 的批量写入路径。"""
    wb = ExcelWorkbook(output_path=tmpfile)
    ws = wb.sheet("Batch")
    rows = [["R1", 1], ["R2", 2], ["R3", 3]]
    wb.write_rows(ws, 1, rows)
    wb.save()
    import os

    assert os.path.exists(tmpfile)
