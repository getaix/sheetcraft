from __future__ import annotations

import json
from typing import Any

from jinja2 import nodes
from jinja2.ext import Extension


class ImageTagExtension(Extension):
    """
    用法：
        {% img path='path/to/img.png', width=100, height=80, fit=true, in_cell=true, keep_ratio=false %}
    说明：
        - path: 图片路径（支持相对模板目录路径，允许动态变量）
        - width/height: 指定尺寸像素
        - fit: 旧逻辑，按单元格宽度适配（可能跨行），优先级低于 in_cell
        - in_cell: 新增，按单元格内嵌，宽高都以单元格为最大边界
        - keep_ratio: 与 in_cell 搭配，是否保持原始宽高比，默认 true
        - scale_x/scale_y: 保持兼容（xlsxwriter 缩放场景）
    """

    tags = {"img"}

    def __init__(self, environment):
        super().__init__(environment)

    def parse(self, parser):
        lineno = next(parser.stream).lineno
        # 第一个必选表达式：path
        path_expr = parser.parse_expression()
        # 可选的命名参数
        kwargs = []
        while parser.stream.current.type == "name":
            name = parser.stream.current.value
            next(parser.stream)  # consume name
            parser.stream.expect("assign")  # '='
            value_expr = parser.parse_expression()
            kwargs.append(nodes.Keyword(name, value_expr))
        call = self.call_method(
            "_emit_placeholder",
            args=[path_expr],
            kwargs=kwargs,
        )
        return nodes.Output([call]).set_lineno(lineno)

    # 运行期：将参数序列化为占位符字符串（保留动态值）
    def _emit_placeholder(self, path, **kwargs):
        import json

        safe_path = "" if path is None else str(path)
        payload = {"path": safe_path}
        for k in (
            "width",
            "height",
            "scale_x",
            "scale_y",
            "fit",
            "in_cell",
            "keep_ratio",
        ):
            if k in kwargs and kwargs[k] is not None:
                payload[k] = kwargs[k]
        return "__FEXCEL_IMG__" + json.dumps(payload, ensure_ascii=False)
