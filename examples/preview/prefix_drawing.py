#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Post-process an .xlsx file to add explicit `xdr:` prefixes to spreadsheetDrawing anchors
(twoCellAnchor / oneCellAnchor / absoluteAnchor) within `xl/drawings/drawing*.xml`.
Some viewers incorrectly expect prefixed tags and fail when the file uses the default namespace.

Usage:
  python prefix_drawing.py input.xlsx output.xlsx
"""
import sys
import zipfile
import re
from io import BytesIO

NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'

ANCHOR_OPEN_RE = re.compile(r'<\s*(twoCellAnchor|oneCellAnchor|absoluteAnchor)([^>]*)>', re.S)
ANCHOR_CLOSE_RE = re.compile(r'</\s*(twoCellAnchor|oneCellAnchor|absoluteAnchor)\s*>', re.S)
ROOT_WS_DR_OPEN_RE = re.compile(r'<\s*wsDr([^>]*)>', re.S)
ROOT_WS_DR_CLOSE_RE = re.compile(r'</\s*wsDr\s*>', re.S)


def add_xdr_prefix(xml: str) -> str:
    changed = False
    # Ensure root has xdr namespace and optionally rename to xdr:wsDr (not strictly required)
    def add_ns_attr(m):
        attrs = m.group(1)
        if 'xmlns:xdr' not in attrs:
            attrs = attrs.strip()
            if attrs.endswith('>'):
                attrs = attrs[:-1]
            attrs += f' xmlns:xdr="{NS_XDR}"'
            changed_local = True
        else:
            changed_local = False
        return f'<wsDr{" " + attrs if attrs else ""}>' , changed_local

    # Add xmlns:xdr attribute if missing
    m = ROOT_WS_DR_OPEN_RE.search(xml)
    if m:
        repl, c = add_ns_attr(m)
        xml = xml[:m.start()] + repl + xml[m.end():]
        changed = changed or c
    
    # Prefix anchors
    def prefix_open(m):
        nonlocal changed
        changed = True
        return f'<xdr:{m.group(1)}{m.group(2)}>'

    def prefix_close(m):
        nonlocal changed
        changed = True
        return f'</xdr:{m.group(1)}>'

    xml = ANCHOR_OPEN_RE.sub(prefix_open, xml)
    xml = ANCHOR_CLOSE_RE.sub(prefix_close, xml)

    return xml, changed


def process_xlsx(input_path: str, output_path: str):
    with zipfile.ZipFile(input_path, 'r') as zin:
        # Create in-memory zip before writing to disk
        bio = BytesIO()
        with zipfile.ZipFile(bio, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for name in zin.namelist():
                data = zin.read(name)
                if name.startswith('xl/drawings/drawing') and name.endswith('.xml'):
                    xml = data.decode('utf-8', errors='ignore')
                    xml2, changed = add_xdr_prefix(xml)
                    if changed:
                        print(f'[prefix] {name}: anchors prefixed')
                        data = xml2.encode('utf-8')
                    else:
                        print(f'[skip] {name}: no change')
                zout.writestr(name, data)
        with open(output_path, 'wb') as f:
            f.write(bio.getvalue())
    print('done ->', output_path)


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print('Usage: python prefix_drawing.py input.xlsx output.xlsx')
        sys.exit(1)
    process_xlsx(sys.argv[1], sys.argv[2])