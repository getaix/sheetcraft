from sheetcraft import ExcelTemplate, ExcelWorkbook


def build_simple_template(path: str):
    wb = ExcelWorkbook(output_path=path, fast=False)
    ws = wb.get_sheet()
    wb.write_cell(ws, 1, 1, "{{ title }}")
    # 使用 Jinja2 原生 for 语法（多行块，容量=2）
    wb.write_cell(ws, 3, 1, "{% for item in items %}")
    wb.write_row(
        ws, 4, ["{{ item.name }}", "{{ item.qty }}", "{{ item.price }}", "=B4*C4"]
    )
    wb.write_row(
        ws, 5, ["{{ item.name }}", "{{ item.qty }}", "{{ item.price }}", "=B5*C5"]
    )
    wb.write_cell(ws, 6, 1, "{% endfor %}")
    wb.save()


def main():
    # Create a simple template workbook
    build_simple_template("template.xlsx")

    data = {
        "title": "Sales Report",
        "items": [
            {"name": "Widget A", "qty": 5, "price": 19.99},
            {"name": "Widget B", "qty": 2, "price": 29.50},
        ],
    }

    # 使用 Jinja2 能力进行模板渲染
    renderer = ExcelTemplate()
    renderer.render("template.xlsx", data, "rendered.xlsx")


if __name__ == "__main__":
    main()
