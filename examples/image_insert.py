from sheetcraft import ExcelWorkbook


def main():
    wb = ExcelWorkbook(output_path="images.xlsx", fast=False)
    ws = wb.add_sheet("Images")
    wb.write_cell(ws, 1, 1, "PNG:")
    wb.insert_image(ws, 2, 1, "example.png", width=160, height=120)

    wb.write_cell(ws, 10, 1, "JPG:")
    wb.insert_image(ws, 11, 1, "example.jpg", width=160, height=120)

    wb.save()


if __name__ == "__main__":
    main()
