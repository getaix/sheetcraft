#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
示例：按 examples/templates 目录中的模板与数据进行渲染，并插入签名图片。

- 模板：到货验收单.xlsx
- 数据：data.json
- 签名图片：img.png（通过 data['owner_signature_id'] 指定）

渲染规则：
- 单元格中的 {{ var }} 会被替换为数据值。
- 使用标准 Jinja2 `{% for %}` 在模板中实现循环块。

图片插入规则（示例约定）：
- 渲染后查找单元格值等于 data['owner_signature_id'] 的位置，将该文本清空并在该单元格锚点插入图片。
  这样模板中只需放置 {{ owner_signature_id }} 占位符即可。
"""

import json
from pathlib import Path
from sheetcraft import ExcelTemplate, FormatFixConfig


def render_acceptance():
    base = Path(__file__).parent
    template_path = base / "到货验收单.xlsx"
    data_path = base / "data.json"
    output_path = base / "到货验收单-渲染结果.xlsx"

    # 加载数据
    with open(data_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # 渲染模板（启用保存后格式修复，以保证 drawings anchors 兼容）
    ExcelTemplate(
        apply_format_fix=True,
        format_fix_config=FormatFixConfig()
    ).render(str(template_path), data, str(output_path))

    print(f"渲染完成：{output_path}")


if __name__ == "__main__":
    render_acceptance()
