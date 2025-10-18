from sheetcraft import ExcelWorkbook


def main():
    wb = ExcelWorkbook(output_path="multi.xlsx", fast=False)
    ws1 = wb.add_sheet("Summary")
    ws2 = wb.add_sheet("Data")
    wb.write_cell(ws1, 1, 1, "Overview")
    wb.write_rows(ws2, 1, [[1, 2, 3], [4, 5, 6], [7, 8, 9]])
    wb.save()


if __name__ == "__main__":
    main()
