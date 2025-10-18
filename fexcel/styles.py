from __future__ import annotations

from typing import Any, Dict, Optional


StyleDict = Dict[str, Any]


def _normalize_color(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    v = value.strip()
    if v.startswith("#"):
        v = v[1:]
    # openpyxl 使用不带 '#' 的 RGB 颜色值
    return v


# region openpyxl
def apply_openpyxl_style(cell, style: Optional[StyleDict]) -> None:
    """为 openpyxl 单元格应用样式。

    参数：
    - `cell`: openpyxl 的单元格对象。
    - `style`: 样式字典，包含 `font`、`align`、`fill`、`border`、`number_format` 等。

    返回：
    - 无。
    """
    if not style:
        return
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    font = style.get("font", {})
    align = style.get("align", {})
    fill = style.get("fill", {})
    border = style.get("border", {})
    num_format = style.get("number_format")

    if font:
        cell.font = Font(
            name=font.get("name"),
            size=font.get("size"),
            bold=font.get("bold"),
            italic=font.get("italic"),
            underline="single" if font.get("underline") else None,
            color=_normalize_color(font.get("color")),
        )
    if align:
        cell.alignment = Alignment(
            horizontal=align.get("horizontal"),
            vertical=align.get("vertical"),
            wrap_text=align.get("wrap"),
        )
    if fill:
        fg = _normalize_color(fill.get("color"))
        if fg:
            cell.fill = PatternFill(fill_type="solid", fgColor=fg)
    if border:

        def _side(val: Optional[Any] = None, color: Optional[str] = None):
            # 支持布尔或具体样式字符串；布尔真值默认使用 "thin"
            style_name = None
            if isinstance(val, bool):
                style_name = "thin" if val else None
            elif isinstance(val, str):
                style_name = val
            return Side(style=style_name, color=_normalize_color(color))

        cell.border = Border(
            left=_side(border.get("left"), border.get("color")),
            right=_side(border.get("right"), border.get("color")),
            top=_side(border.get("top"), border.get("color")),
            bottom=_side(border.get("bottom"), border.get("color")),
        )
    if num_format:
        cell.number_format = num_format


# endregion


# region xlsxwriter
def xlsxwriter_format_from_dict(wb, style: Optional[StyleDict]):
    """将样式字典转换为 xlsxwriter 的 `Format` 对象，并对重复样式进行缓存。

    参数：
    - `wb`: xlsxwriter 的 `Workbook` 对象。
    - `style`: 样式字典。

    返回：
    - xlsxwriter `Format` 对象或 `None`。
    """
    if not style:
        return None

    # 将样式字典映射为 xlsxwriter 的格式字典（与 add_format 兼容）
    def _to_xlsxwriter_dict(style_dict: StyleDict) -> Dict[str, Any]:
        fmt: Dict[str, Any] = {}
        font = style_dict.get("font", {})
        align = style_dict.get("align", {})
        fill = style_dict.get("fill", {})
        border = style_dict.get("border", {})
        num_format = style_dict.get("number_format")

        if font:
            if font.get("bold"):
                fmt["bold"] = True
            if font.get("italic"):
                fmt["italic"] = True
            if font.get("underline"):
                fmt["underline"] = 1
            if font.get("size"):
                fmt["font_size"] = font.get("size")
            if font.get("name"):
                fmt["font_name"] = font.get("name")
            if font.get("color"):
                fmt["font_color"] = font.get("color")
        if align:
            if align.get("horizontal"):
                fmt["align"] = align.get("horizontal")
            if align.get("vertical"):
                fmt["valign"] = align.get("vertical")
            if align.get("wrap"):
                fmt["text_wrap"] = True
        if fill:
            if fill.get("color"):
                fmt["bg_color"] = fill.get("color")
                fmt["pattern"] = 1
        if border:
            # 简单的四周边框
            if any(border.get(k) for k in ("left", "right", "top", "bottom")):
                fmt["border"] = 1
            if border.get("color"):
                fmt["border_color"] = border.get("color")
        if num_format:
            fmt["num_format"] = num_format
        return fmt

    fmt = _to_xlsxwriter_dict(style)

    # 对重复样式进行缓存，减少 add_format 次数
    cache = getattr(wb, "_fexcel_format_cache", None)
    if cache is None:
        cache = {}
        setattr(wb, "_fexcel_format_cache", cache)
    key = tuple(sorted(fmt.items()))
    cached = cache.get(key)
    if cached is not None:
        return cached

    fmt_obj = wb.add_format(fmt)
    cache[key] = fmt_obj
    return fmt_obj


# endregion


# region xlwt
def xlwt_style_from_dict(style: Optional[StyleDict]):
    """将样式字典转换为 xlwt 的 `XFStyle` 对象（带轻量缓存）。

    参数：
    - `style`: 样式字典。

    返回：
    - xlwt `XFStyle` 对象或 `None`。
    """
    if not style:
        return None
    import xlwt

    # 轻量缓存：避免重复创建相同样式对象
    cache = getattr(xlwt_style_from_dict, "_cache", None)
    if cache is None:
        cache = {}
        setattr(xlwt_style_from_dict, "_cache", cache)
    # 使用排序后的键确保相同字典生成一致键
    try:
        key_items = []
        for k, v in style.items():
            if isinstance(v, dict):
                key_items.append((k, tuple(sorted(v.items()))))
            else:
                key_items.append((k, v))
        key = tuple(sorted(key_items))
    except Exception:
        key = None
    if key is not None and key in cache:
        return cache[key]

    xf = xlwt.XFStyle()

    font = style.get("font", {})
    align = style.get("align", {})
    fill = style.get("fill", {})
    border = style.get("border", {})

    if font:
        f = xlwt.Font()
        if font.get("bold"):
            f.bold = True
        if font.get("italic"):
            f.italic = True
        if font.get("underline"):
            f.underline = xlwt.Font.UNDERLINE_SINGLE
        if font.get("size"):
            f.height = int(font.get("size") * 20)
        if font.get("name"):
            f.name = font.get("name")
        xf.font = f

    if align:
        a = xlwt.Alignment()
        horiz = align.get("horizontal")
        vert = align.get("vertical")
        if horiz:
            mapping = {
                "left": xlwt.Alignment.HORZ_LEFT,
                "center": xlwt.Alignment.HORZ_CENTER,
                "right": xlwt.Alignment.HORZ_RIGHT,
            }
            a.horz = mapping.get(horiz, xlwt.Alignment.HORZ_GENERAL)
        if vert:
            mapping = {
                "top": xlwt.Alignment.VERT_TOP,
                "center": xlwt.Alignment.VERT_CENTER,
                "bottom": xlwt.Alignment.VERT_BOTTOM,
            }
            a.vert = mapping.get(vert, xlwt.Alignment.VERT_CENTER)
        if align.get("wrap"):
            a.wrap = 1
        xf.alignment = a

    if fill:
        p = xlwt.Pattern()
        if fill.get("color"):
            p.pattern = xlwt.Pattern.SOLID_PATTERN
            # xlwt 使用调色板索引；保持默认调色板
        xf.pattern = p

    if border:
        b = xlwt.Borders()
        # 基本细边框
        if border.get("left"):
            b.left = xlwt.Borders.THIN
        if border.get("right"):
            b.right = xlwt.Borders.THIN
        if border.get("top"):
            b.top = xlwt.Borders.THIN
        if border.get("bottom"):
            b.bottom = xlwt.Borders.THIN
        xf.borders = b

    # xlwt 的数字格式能力有限；此处略过
    if key is not None:
        cache[key] = xf
    return xf


# endregion
