from sheetcraft import ExcelWorkbook
from sheetcraft.workbook import DataValidationSpec


def main():
    wb = ExcelWorkbook(output_path="out.xlsx", fast=True)
    ws = wb.add_sheet("Report")

    header_style = {
        "font": {"bold": True, "size": 12},
        "fill": {"color": "#DDEEFF"},
        "align": {"horizontal": "center"},
        "border": {"left": True, "right": True, "top": True, "bottom": True},
    }
    wb.write_row(ws, 1, ["Item", "Qty", "Price", "Total"], styles=[header_style] * 4)

    rows = [
        ["Widget A", 5, 19.99, "=B2*C2"],
        ["Widget B", 2, 29.50, "=B3*C3"],
    ]
    wb.write_rows(ws, start_row=2, rows=rows)

    wb.add_data_validation(
        ws, "B2:B100", DataValidationSpec(type="whole", operator=">=", formula1="0")
    )
    wb.insert_image(ws, 1, 6, "logo.png", scale_x=0.5, scale_y=0.5)

    wb.set_column_width(ws, 1, 18)
    wb.set_column_width(ws, 2, 10)
    wb.set_column_width(ws, 3, 12)
    wb.set_column_width(ws, 4, 12)

    wb.save()


if __name__ == "__main__":
    main()
