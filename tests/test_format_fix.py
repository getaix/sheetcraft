import zipfile
from pathlib import Path

from fexcel.format_fix import _add_xdr_prefix, _normalize_sheet_rels, fix_xlsx, FormatFixConfig


def test_add_xdr_prefix_on_wsdr_and_anchors():
    xml = (
        "<wsDr>"
        "<twoCellAnchor><from/></twoCellAnchor>"
        "<oneCellAnchor><from/></oneCellAnchor>"
        "</wsDr>"
    )
    new_xml, changed, logs = _add_xdr_prefix(xml)
    assert changed is True
    assert "xdr:wsDr" in new_xml and "xmlns:xdr" in new_xml
    assert "<xdr:twoCellAnchor" in new_xml and "</xdr:twoCellAnchor>" in new_xml
    assert any("anchors" in msg or "wsDr" in msg for msg in logs)


def test_normalize_sheet_rels_targets():
    xml = (
        "<Relationships>"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"/xl/drawings/drawing1.xml\"/>"
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"xl/drawings/drawing2.xml\"/>"
        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"drawings/drawing3.xml\"/>"
        "<Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing4.xml\"/>"
        "</Relationships>"
    )
    new_xml, changed, logs = _normalize_sheet_rels(xml)
    assert changed is True
    # Ensure normalization to ../drawings/ for first three entries
    assert "Target=\"../drawings/drawing1.xml\"" in new_xml
    assert "Target=\"../drawings/drawing2.xml\"" in new_xml
    assert "Target=\"../drawings/drawing3.xml\"" in new_xml
    # Already normalized stays unchanged
    assert "Target=\"../drawings/drawing4.xml\"" in new_xml
    assert any("rels:" in msg for msg in logs)


def test_fix_xlsx_end_to_end(tmp_path):
    # Create a minimal zip shaped like xlsx with drawing and rels
    src = tmp_path / "src.xlsx"
    with zipfile.ZipFile(src, "w") as z:
        z.writestr(
            "xl/drawings/drawing1.xml",
            "<wsDr><twoCellAnchor><from/></twoCellAnchor></wsDr>",
        )
        z.writestr(
            "xl/worksheets/_rels/sheet1.xml.rels",
            "<Relationships>"
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"/xl/drawings/drawing1.xml\"/>"
            "</Relationships>",
        )
    out = tmp_path / "out.xlsx"
    cfg = FormatFixConfig(prefix_drawing_anchors=True, normalize_sheet_rel_targets=True)
    report = fix_xlsx(str(src), str(out), cfg)
    assert out.exists()
    # Verify report records
    assert "prefix_drawing_anchors" in report.rules_applied
    assert "normalize_sheet_rel_targets" in report.rules_applied
    assert any(name.endswith("drawing1.xml") for name in report.changed_entries)
    assert any(name.endswith("sheet1.xml.rels") for name in report.changed_entries)
    # Verify content changes
    with zipfile.ZipFile(out, "r") as z:
        dxml = z.read("xl/drawings/drawing1.xml").decode("utf-8")
        rels = z.read("xl/worksheets/_rels/sheet1.xml.rels").decode("utf-8")
    assert "<xdr:wsDr" in dxml and "xmlns:xdr" in dxml
    assert "<xdr:twoCellAnchor" in dxml and "</xdr:twoCellAnchor>" in dxml
    assert "Target=\"../drawings/drawing1.xml\"" in rels


from fexcel.format_fix import (
    _ensure_xml_decl,
    _normalize_drawing_rels,
    _normalize_workbook_rels,
    _reorder_relationship_attrs,
    _reorder_wsdr_xmlns,
)


def test_reorder_wsdr_xmlns_and_xml_decl():
    xml = "<xdr:wsDr xmlns:r=\"rns\" xmlns:a=\"ans\" xmlns:xdr=\"xns\" other=\"1\"></xdr:wsDr>"
    new_xml, changed, logs = _reorder_wsdr_xmlns(xml)
    assert changed is True
    assert new_xml.index("xmlns:xdr") < new_xml.index("xmlns:r") < new_xml.index("xmlns:a")
    decl_xml, c2, l2 = _ensure_xml_decl(new_xml)
    assert c2 is True
    assert decl_xml.startswith("<?xml ")


def test_rels_normalizers_and_attr_order_extended():
    # drawing rels normalization
    drawing_xml = (
        "<Relationships>\n"
        "<Relationship Target=\"xl/media/image1.png\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Id=\"rId2\"/>"
        "</Relationships>"
    )
    new_drawing, changed_drawing, _ = _normalize_drawing_rels(drawing_xml)
    assert changed_drawing is True
    assert "../media/image1.png" in new_drawing
    # workbook rels normalization
    wb_xml = (
        "<Relationships>\n"
        "<Relationship Target=\"/xl/worksheets/sheet1.xml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Id=\"rId3\"/>"
        "</Relationships>"
    )
    new_wb, changed_wb, _ = _normalize_workbook_rels(wb_xml)
    assert changed_wb is True
    assert "worksheets/sheet1.xml" in new_wb
    # reorder relationship attributes
    sheet_xml = (
        "<Relationships>\n"
        "<Relationship Target=\"/xl/drawings/drawing1.xml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Id=\"rId1\"/>"
        "</Relationships>"
    )
    ordered_sheet, ordered_changed, logs = _reorder_relationship_attrs(sheet_xml)
    assert ordered_changed is True
    assert any("属性顺序" in s for s in logs)
    import re as _re
    m = _re.search(r"<Relationship ([^>]*)/?>", ordered_sheet)
    assert m
    attrs = m.group(1)
    assert attrs.index("Id=") < attrs.index("Type=") < attrs.index("Target=")