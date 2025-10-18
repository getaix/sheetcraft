from __future__ import annotations


def add_data_validation_openpyxl(
    ws, cell_range: str, spec: "DataValidationSpec"
) -> None:
    """在 openpyxl 工作表中为指定范围添加数据有效性规则。

    参数：
    - `ws`: openpyxl 工作表对象。
    - `cell_range`: 单元格范围字符串（如 `A1:D10`）。
    - `spec`: 数据有效性规则规格，参见 `DataValidationSpec`。

    返回：
    - 无。
    """
    from openpyxl.worksheet.datavalidation import DataValidation

    # 操作符映射到 openpyxl 的合法枚举
    op_map = {
        ">": "greaterThan",
        ">=": "greaterThanOrEqual",
        "<": "lessThan",
        "<=": "lessThanOrEqual",
        "==": "equal",
        "!=": "notEqual",
        "between": "between",
        "notBetween": "notBetween",
    }
    operator = op_map.get(spec.operator, spec.operator)

    dv = DataValidation(
        type=spec.type,
        formula1=spec.formula1,
        formula2=spec.formula2,
        operator=operator,
        allow_blank=spec.allow_blank,
        showInputMessage=spec.show_input_message,
        promptTitle=spec.prompt_title,
        prompt=spec.prompt,
        showErrorMessage=spec.show_error_message,
        errorTitle=spec.error_title,
        error=spec.error,
    )
    ws.add_data_validation(dv)
    dv.add(cell_range)


def add_data_validation_xlsxwriter(
    ws,
    first_row: int,
    first_col: int,
    last_row: int,
    last_col: int,
    spec: "DataValidationSpec",
) -> None:
    """通过 xlsxwriter 的 `data_validation` API 添加数据有效性。

    说明：
    - 本库对行列使用 1 基坐标，这里需转换为 0 基以调用底层 API。

    参数：
    - `ws`: xlsxwriter 工作表对象。
    - `first_row`, `first_col`, `last_row`, `last_col`: 范围的 1 基行列边界。
    - `spec`: 数据有效性规格。

    返回：
    - 无。
    """

    options = {
        "validate": spec.type,
        "criteria": spec.operator if spec.operator else None,
        "source": None,
        "value": None,
        "minimum": None,
        "maximum": None,
        "ignore_blank": spec.allow_blank,
        "input_title": spec.prompt_title,
        "input_message": spec.prompt,
        "error_title": spec.error_title,
        "error_message": spec.error,
    }

    # Map formulas depending on type
    if spec.type == "list":
        # For list, use 'source' which can be range or CSV string
        options["source"] = spec.formula1
    elif spec.type in {
        "whole",
        "decimal",
        "date",
        "time",
        "text_length",
    }:  # xlsxwriter names
        # numeric/date types commonly use minimum/maximum with criteria 'between'/'not between'
        if spec.operator in {"between", "not between", "between", "notBetween"}:
            options["minimum"] = spec.formula1
            options["maximum"] = spec.formula2
        else:
            # single value comparison
            options["value"] = spec.formula1
    elif spec.type == "custom":
        options["value"] = spec.formula1

    # Clean None values to avoid xlsxwriter errors
    options = {k: v for k, v in options.items() if v is not None}

    ws.data_validation(
        first_row - 1, first_col - 1, last_row - 1, last_col - 1, options
    )
