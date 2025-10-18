from __future__ import annotations

"""
面向 `.xlsx` 的专用、解耦合导出后格式修复模块。

设计目标：
- 与核心业务逻辑解耦合：仅调整文件结构/标记，不改变导出数据。
- 保持原始数据完整性：不触及单元格值、公式或样式。
- 提供清晰日志/报告：每条修复动作均可追踪。
- 支持可配置规则：按需开启或关闭具体修复。

当前规则：
- prefix_drawing_anchors：为 `xl/drawings/drawing*.xml` 中的 spreadsheetDrawing anchors
 （twoCellAnchor/oneCellAnchor/absoluteAnchor）显式加上 `xdr:` 前缀；
  若根元素 `wsDr` 缺少 `xmlns:xdr` 则追加该命名空间属性，并将根标签统一为 `xdr:wsDr`。
- normalize_sheet_rel_targets：规范 `xl/worksheets/_rels/sheet*.xml.rels` 中 Drawing 关系的 `Target` 为相对路径
 （例如将 `/xl/drawings/drawing1.xml` 或 `xl/drawings/drawing1.xml` 统一为 `../drawings/drawing1.xml`）。
- normalize_drawing_rel_targets：规范 `xl/drawings/_rels/drawing*.xml.rels` 中 Image 关系的 `Target` 为相对路径
 （例如将 `/xl/media/image1.png` 或 `xl/media/image1.png` 统一为 `../media/image1.png`）。

模块通过 `zipfile` 在 `.xlsx` 压缩包层面工作：读取条目、按需改写少量文本条目，
并写回到原文件或指定输出路径。
"""

import re
import zipfile
from dataclasses import dataclass, field
from io import BytesIO
from typing import List, Optional, Tuple

XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"

# 预编译正则：提高性能与匹配稳定性
ANCHOR_OPEN_RE = re.compile(r"<\s*(twoCellAnchor|oneCellAnchor|absoluteAnchor)([^>]*)>", re.S)
ANCHOR_CLOSE_RE = re.compile(r"</\s*(twoCellAnchor|oneCellAnchor|absoluteAnchor)\s*>", re.S)
# 同时匹配未带前缀与已带前缀的 wsDr 根标签
WS_DR_OPEN_RE = re.compile(r"<\s*(?:xdr:)?wsDr([^>]*)>", re.S)
WS_DR_CLOSE_RE = re.compile(r"</\s*(?:xdr:)?wsDr\s*>", re.S)
# 新增：匹配所有未带命名空间前缀且应属于 xdr 命名空间的标签（避免影响 a:/r:/xdr: 已带前缀的标签）
GEN_XDR_OPEN_RE = re.compile(r"<\s*(?!/?(?:xdr:|a:|r:))([A-Za-z][A-Za-z0-9_]*)([^>]*)>", re.S)
GEN_XDR_CLOSE_RE = re.compile(r"</\s*(?!(?:xdr:|a:|r:))([A-Za-z][A-Za-z0-9_]*)\s*>", re.S)

# 匹配 worksheet 关系标签中针对 drawing 的 Target（支持属性顺序变化和单双引号）
REL_DRAW_TARGET_RE = re.compile(
    r"(<Relationship\b(?=[^>]*\bType\s*=\s*[\"']http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing[\"'])"
    r"(?:[^>]*?)\bTarget\s*=\s*[\"'])([^\"']+)([\"'])",
    re.S,
)

# 匹配 drawing 关系标签中针对 image 的 Target
REL_IMAGE_TARGET_RE = re.compile(
    r"(<Relationship\b(?=[^>]*\bType\s*=\s*[\"']http://schemas.openxmlformats.org/officeDocument/2006/relationships/image[\"'])"
    r"(?:[^>]*?)\bTarget\s*=\s*[\"'])([^\"']+)([\"'])",
    re.S,
)

# 匹配 workbook 关系文件中 worksheet 的 Target
REL_WORKSHEET_TARGET_RE = re.compile(
    r"(<Relationship\b(?=[^>]*\bType\s*=\s*[\"']http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet[\"'])"
    r"(?:[^>]*?)\bTarget\s*=\s*[\"'])([^\"']+)([\"'])",
    re.S,
)
# 匹配任意 Relationship 标签（用于统一属性顺序）
REL_WORKSHEET_TAG_RE = re.compile(r"<Relationship\b([^>]*)/?\s*>", re.S)


@dataclass
class FormatFixConfig:
    """格式修复配置。

    - prefix_drawing_anchors: 是否为锚点标签添加 `xdr:` 前缀。
    - normalize_sheet_rel_targets: 是否规范化 worksheet 关系文件中的 Drawing `Target` 为相对路径。
    - normalize_drawing_rel_targets: 是否规范化 drawing 关系文件中的 Image `Target` 为相对路径。
    - normalize_workbook_rel_targets: 是否规范化 workbook 关系文件中的 Worksheet `Target` 为相对路径。
    - ensure_xml_decl: 是否为目标 XML/RELS 文件补充 XML 声明头。
    - reorder_wsdr_xmlns: 是否规范 `xdr:wsDr` 根的命名空间属性顺序为 xdr、r、a。
    - reorder_relationship_attrs: 是否统一 Relationship 标签属性顺序为 Id、Type、Target。
    """

    prefix_drawing_anchors: bool = True
    normalize_sheet_rel_targets: bool = True
    normalize_drawing_rel_targets: bool = True
    normalize_workbook_rel_targets: bool = True
    ensure_xml_decl: bool = True
    reorder_wsdr_xmlns: bool = True
    reorder_relationship_attrs: bool = True
    # 预留后续规则（保持接口稳定）
    # normalize_content_types: bool = False
    # sanitize_shared_strings: bool = False


@dataclass
class FixReport:
    """格式修复操作的结果报告。"""

    input_path: str
    output_path: str
    in_place: bool
    rules_applied: List[str] = field(default_factory=list)
    changed_entries: List[str] = field(default_factory=list)
    logs: List[str] = field(default_factory=list)

    def log(self, message: str) -> None:
        self.logs.append(message)


def _add_xdr_prefix(xml: str) -> Tuple[str, bool, List[str]]:
    """返回三元组 `(new_xml, changed, logs)`，为绘图部分的 xdr 命名空间进行规范化。

    - 统一根 `wsDr` 为 `xdr:wsDr`，并确保存在 `xmlns:xdr`。
    - 移除根上的默认 `xmlns`（spreadsheatDrawing）以避免与前缀并存造成歧义。
    - 为锚点标签（twoCellAnchor/oneCellAnchor/absoluteAnchor）的开闭标签添加 `xdr:` 前缀。
    - 为其他未带前缀且应属于 xdr 命名空间的标签（如 from/to/pic/nvPicPr/blipFill/spPr/clientData 等）补全 `xdr:` 前缀。
    """
    logs: List[str] = []
    changed = False

    # 统一根标签为 xdr:wsDr，并确保包含 xmlns:xdr 命名空间属性，同时移除默认 xmlns
    m = WS_DR_OPEN_RE.search(xml)
    if m:
        attrs = m.group(1)
        attrs_clean = attrs.strip()
        # 移除默认 xmlns
        attrs_clean_no_default = re.sub(
            rf"\s+xmlns=\"{re.escape(XDR_NS)}\"",
            "",
            attrs_clean,
        )
        has_ns = "xmlns:xdr" in attrs_clean_no_default
        # 保证在标签名与属性之间存在空格
        attrs_part = f" {attrs_clean_no_default}" if attrs_clean_no_default else ""
        if not has_ns:
            attrs_part = f"{attrs_part} xmlns:xdr=\"{XDR_NS}\""
        new_open = f"<xdr:wsDr{attrs_part}>"
        xml = xml[: m.start()] + new_open + xml[m.end() :]
        changed = True
        logs.append("wsDr: 已统一为 xdr:wsDr，移除默认 xmlns 并确保 xmlns:xdr 存在")
        # 同步关闭标签
        xml = WS_DR_CLOSE_RE.sub("</xdr:wsDr>", xml)

    # 锚点标签的开闭标签加前缀
    def prefix_open(m):
        nonlocal changed
        changed = True
        tag, rest = m.group(1), m.group(2)
        return f"<xdr:{tag}{rest}>"

    def prefix_close(m):
        nonlocal changed
        changed = True
        tag = m.group(1)
        return f"</xdr:{tag}>"

    new_xml = ANCHOR_OPEN_RE.sub(prefix_open, xml)
    new_xml = ANCHOR_CLOSE_RE.sub(prefix_close, new_xml)

    # 其他未带前缀的 xdr 元素统一补前缀（不影响 a:/r:/xdr: 已带前缀的标签）
    new_xml2 = GEN_XDR_OPEN_RE.sub(lambda m: f"<xdr:{m.group(1)}{m.group(2)}>", new_xml)
    new_xml2 = GEN_XDR_CLOSE_RE.sub(lambda m: f"</xdr:{m.group(1)}>", new_xml2)

    if new_xml2 != xml:
        changed = True
        # 根据变化情况记录日志
        if new_xml2 != new_xml:
            logs.append("elements: 已为未带前缀的 xdr 元素补全前缀")
        else:
            logs.append("anchors: 已为锚点标签添加 xdr 前缀")

    return new_xml2, changed, logs


def _normalize_target_path_for_sheet(target: str) -> str:
    """将绘图关系的 Target 规范化为相对源部件（`xl/worksheets/sheetN.xml`）的路径。

    - 去除前导 `/`（绝对路径）。
    - 去除前缀 `xl/`。
    - 将 `drawings/...` 统一前置 `../`，得到 `../drawings/...`。
    - 若已是 `../drawings/...`，保持不变。
    """
    t = target.strip()
    if t.startswith("/"):
        t = t[1:]
    if t.startswith("xl/"):
        t = t[3:]
    if t.startswith("../drawings/"):
        return t
    if t.startswith("drawings/"):
        return "../" + t
    # 其他情况保持原样
    return t


def _normalize_target_path_for_drawing(target: str) -> str:
    """将图片关系的 Target 规范化为相对源部件（`xl/drawings/drawingN.xml`）的路径。

    - 去除前导 `/`（绝对路径）。
    - 去除前缀 `xl/`。
    - 将 `media/...` 统一前置 `../`，得到 `../media/...`。
    - 若已是 `../media/...`，保持不变。
    """
    t = target.strip()
    if t.startswith("/"):
        t = t[1:]
    if t.startswith("xl/"):
        t = t[3:]
    if t.startswith("../media/"):
        return t
    if t.startswith("media/"):
        return "../" + t
    # 其他情况保持原样
    return t


def _normalize_target_path_for_workbook(target: str) -> str:
    """将 workbook 关系中的 worksheet Target 规范化为相对路径（相对于 `xl/workbook.xml`）。

    - 去除前导 `/`（绝对路径）。
    - 去除前缀 `xl/`。
    - 保持为 `worksheets/...` 形式。
    """
    t = target.strip()
    if t.startswith("/"):
        t = t[1:]
    if t.startswith("xl/"):
        t = t[3:]
    return t


def _normalize_sheet_rels(xml: str) -> Tuple[str, bool, List[str]]:
    """规范 worksheet 关系文件中的 drawing Target 路径。

    返回 `(new_xml, changed, logs)`。
    """
    changed = False
    logs: List[str] = []

    def repl(m: re.Match) -> str:
        nonlocal changed, logs
        prefix, target, quote = m.group(1), m.group(2), m.group(3)
        new_target = _normalize_target_path_for_sheet(target)
        if new_target != target:
            changed = True
            logs.append(f"rels: drawing Target 修正 {target} -> {new_target}")
        return f"{prefix}{new_target}{quote}"

    new_xml = REL_DRAW_TARGET_RE.sub(repl, xml)
    return new_xml, changed, logs


def _normalize_drawing_rels(xml: str) -> Tuple[str, bool, List[str]]:
    """规范 drawing 关系文件中的 image Target 路径。

    返回 `(new_xml, changed, logs)`。
    """
    changed = False
    logs: List[str] = []

    def repl(m: re.Match) -> str:
        nonlocal changed, logs
        prefix, target, quote = m.group(1), m.group(2), m.group(3)
        new_target = _normalize_target_path_for_drawing(target)
        if new_target != target:
            changed = True
            logs.append(f"rels: image Target 修正 {target} -> {new_target}")
        return f"{prefix}{new_target}{quote}"

    new_xml = REL_IMAGE_TARGET_RE.sub(repl, xml)
    return new_xml, changed, logs


def _normalize_workbook_rels(xml: str) -> Tuple[str, bool, List[str]]:
    """规范 workbook 关系文件中的 worksheet Target 路径为相对路径。"""
    changed = False
    logs: List[str] = []

    def repl(m: re.Match) -> str:
        nonlocal changed, logs
        prefix, target, quote = m.group(1), m.group(2), m.group(3)
        new_target = _normalize_target_path_for_workbook(target)
        if new_target != target:
            changed = True
            logs.append(f"rels: worksheet Target 修正 {target} -> {new_target}")
        return f"{prefix}{new_target}{quote}"

    new_xml = REL_WORKSHEET_TARGET_RE.sub(repl, xml)
    return new_xml, changed, logs


def _ensure_xml_decl(xml: str) -> Tuple[str, bool, List[str]]:
    if xml.lstrip().startswith("<?xml"):
        return xml, False, []
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" + xml, True, ["xml: 添加 XML 声明头"]


def _reorder_wsdr_xmlns(xml: str) -> Tuple[str, bool, List[str]]:
    m = WS_DR_OPEN_RE.search(xml)
    if not m:
        return xml, False, []
    attrs = m.group(1)
    orig = attrs
    # 提取并移除命名空间属性
    ns_vals = {}
    for key in ("xdr", "r", "a"):
        xmln_re = re.compile(rf"\s+xmlns:{key}=\"([^\"]+)\"")
        mm = xmln_re.search(attrs)
        if mm:
            ns_vals[key] = mm.group(1)
            attrs = xmln_re.sub("", attrs)
    # 重新拼接为指定顺序
    new_attrs = ""
    for key in ("xdr", "r", "a"):
        if key in ns_vals:
            new_attrs += f" xmlns:{key}=\"{ns_vals[key]}\""
    attrs_rest = attrs.strip()
    if attrs_rest:
        new_attrs += f" {attrs_rest}"
    if new_attrs == f" {orig.strip()}" or new_attrs.strip() == orig.strip():
        return xml, False, []
    new_open = f"<xdr:wsDr{new_attrs}>"
    new_xml = xml[:m.start()] + new_open + xml[m.end():]
    return new_xml, True, ["wsDr: 规范命名空间属性顺序为 xdr、r、a"]


def _reorder_relationship_attrs(xml: str) -> Tuple[str, bool, List[str]]:
    changed = False
    logs: List[str] = []

    def rebuild_tag(m: re.Match) -> str:
        nonlocal changed
        attrs = m.group(1)
        # 解析属性对
        pairs = re.findall(r"(\w+)\s*=\s*['\"]([^'\"]+)['\"]", attrs)
        if not pairs:
            return m.group(0)
        values = {}
        order_rest = []
        for k, v in pairs:
            if k in ("Id", "Type", "Target"):
                values[k] = v
            else:
                order_rest.append((k, v))
        # 构建新顺序：Id, Type, Target, 其他按原顺序
        buf = ""
        for k in ("Id", "Type", "Target"):
            if k in values:
                buf += f" {k}=\"{values[k]}\""
        for k, v in order_rest:
            buf += f" {k}=\"{v}\""
        orig_attrs_norm = " ".join([f"{k}=\"{v}\"" for k, v in pairs])
        new_attrs_norm = buf.strip()
        if new_attrs_norm != orig_attrs_norm:
            changed = True
        closing = "/>" if m.group(0).strip().endswith("/>") else ">"
        # 保持自闭合时的空格风格
        spacer = ""
        return f"<Relationship{buf}{spacer}{closing}"

    new_xml = REL_WORKSHEET_TAG_RE.sub(rebuild_tag, xml)
    if changed:
        logs.append("rels: 统一 Relationship 属性顺序为 Id、Type、Target")
    return new_xml, changed, logs


def fix_xlsx(input_path: str, output_path: Optional[str] = None, config: Optional[FormatFixConfig] = None) -> FixReport:
    """对 `.xlsx` 文件应用配置化的格式修复。

    - input_path: 源 `.xlsx` 路径。
    - output_path: 目标路径；若为 None，则对源文件就地覆盖（in-place）。
    - config: `FormatFixConfig`；默认启用前缀修复。

    返回：`FixReport`，包含详细日志与被改写的条目列表。
    """
    cfg = config or FormatFixConfig()
    out = output_path or input_path
    report = FixReport(input_path=input_path, output_path=out, in_place=output_path is None)

    with zipfile.ZipFile(input_path, "r") as zin:
        buf = BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for name in zin.namelist():
                data = zin.read(name)
                # 1) 修复 drawings anchors/xdr 命名空间
                if cfg.prefix_drawing_anchors and name.startswith("xl/drawings/drawing") and name.endswith(".xml"):
                    try:
                        s = data.decode("utf-8", errors="ignore")
                        s2, changed, logs = _add_xdr_prefix(s)
                        agg_changed = changed
                        agg_logs = list(logs)
                        if cfg.reorder_wsdr_xmlns:
                            s3, c3, l3 = _reorder_wsdr_xmlns(s2)
                            if c3:
                                agg_changed = True
                                agg_logs.extend(l3)
                            s2 = s3
                        if cfg.ensure_xml_decl:
                            s4, c4, l4 = _ensure_xml_decl(s2)
                            if c4:
                                agg_changed = True
                                agg_logs.extend(l4)
                            s2 = s4
                        if agg_changed:
                            data = s2.encode("utf-8")
                            report.changed_entries.append(name)
                            report.rules_applied.append("prefix_drawing_anchors")
                            for msg in agg_logs:
                                report.log(f"[{name}] {msg}")
                        else:
                            report.log(f"[{name}] 无变化")
                    except Exception as exc:
                        report.log(f"[{name}] prefix_drawing_anchors 处理失败: {exc}")
                # 2) 规范 worksheet 关系的 Target 路径
                elif (
                    cfg.normalize_sheet_rel_targets
                    and name.startswith("xl/worksheets/_rels/")
                    and name.endswith(".rels")
                ):
                    try:
                        s = data.decode("utf-8", errors="ignore")
                        s2, changed, logs = _normalize_sheet_rels(s)
                        agg_changed = changed
                        agg_logs = list(logs)
                        if cfg.reorder_relationship_attrs:
                            s3, c3, l3 = _reorder_relationship_attrs(s2)
                            if c3:
                                agg_changed = True
                                agg_logs.extend(l3)
                            s2 = s3
                        if cfg.ensure_xml_decl:
                            s4, c4, l4 = _ensure_xml_decl(s2)
                            if c4:
                                agg_changed = True
                                agg_logs.extend(l4)
                            s2 = s4
                        if agg_changed:
                            data = s2.encode("utf-8")
                            report.changed_entries.append(name)
                            report.rules_applied.append("normalize_sheet_rel_targets")
                            for msg in agg_logs:
                                report.log(f"[{name}] {msg}")
                        else:
                            report.log(f"[{name}] 无变化")
                    except Exception as exc:
                        report.log(f"[{name}] normalize_sheet_rel_targets 处理失败: {exc}")
                # 3) 规范 drawing 关系的 Target 路径（image）
                elif (
                    cfg.normalize_drawing_rel_targets
                    and name.startswith("xl/drawings/_rels/")
                    and name.endswith(".rels")
                ):
                    try:
                        s = data.decode("utf-8", errors="ignore")
                        s2, changed, logs = _normalize_drawing_rels(s)
                        agg_changed = changed
                        agg_logs = list(logs)
                        if cfg.reorder_relationship_attrs:
                            s3, c3, l3 = _reorder_relationship_attrs(s2)
                            if c3:
                                agg_changed = True
                                agg_logs.extend(l3)
                            s2 = s3
                        if cfg.ensure_xml_decl:
                            s4, c4, l4 = _ensure_xml_decl(s2)
                            if c4:
                                agg_changed = True
                                agg_logs.extend(l4)
                            s2 = s4
                        if agg_changed:
                            data = s2.encode("utf-8")
                            report.changed_entries.append(name)
                            report.rules_applied.append("normalize_drawing_rel_targets")
                            for msg in agg_logs:
                                report.log(f"[{name}] {msg}")
                        else:
                            report.log(f"[{name}] 无变化")
                    except Exception as exc:
                        report.log(f"[{name}] normalize_drawing_rel_targets 处理失败: {exc}")
                # 4) 规范 workbook 关系的 Target 路径（worksheet）
                elif cfg.normalize_workbook_rel_targets and name == "xl/_rels/workbook.xml.rels":
                    try:
                        s = data.decode("utf-8", errors="ignore")
                        s2, changed, logs = _normalize_workbook_rels(s)
                        agg_changed = changed
                        agg_logs = list(logs)
                        if cfg.reorder_relationship_attrs:
                            s3, c3, l3 = _reorder_relationship_attrs(s2)
                            if c3:
                                agg_changed = True
                                agg_logs.extend(l3)
                            s2 = s3
                        if cfg.ensure_xml_decl:
                            s4, c4, l4 = _ensure_xml_decl(s2)
                            if c4:
                                agg_changed = True
                                agg_logs.extend(l4)
                            s2 = s4
                        if agg_changed:
                            data = s2.encode("utf-8")
                            report.changed_entries.append(name)
                            report.rules_applied.append("normalize_workbook_rel_targets")
                            for msg in agg_logs:
                                report.log(f"[{name}] {msg}")
                        else:
                            report.log(f"[{name}] 无变化")
                    except Exception as exc:
                        report.log(f"[{name}] normalize_workbook_rel_targets 处理失败: {exc}")

                zout.writestr(name, data)
        # 写入目标
        with open(out, "wb") as f:
            f.write(buf.getvalue())

    report.log(f"saved -> {out} (in_place={report.in_place})")
    return report