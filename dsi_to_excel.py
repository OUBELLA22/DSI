#!/usr/bin/env python3
"""
Convert a Capital HarnessXC DSI export into a multi-sheet, formatted .xlsx workbook.

No third-party dependencies: an .xlsx file is a ZIP of XML parts, which we build
directly with the standard library. Each DSI section becomes its own worksheet,
rendered as a native Excel Table (banded rows, filter, bold header) with frozen
header row and auto-sized columns.
"""

import re
import sys
import zipfile
from datetime import datetime
from html import escape

DELIM = ":"


# --------------------------------------------------------------------------- #
# Section schemas: meaningful headers derived from Capital HarnessXC DSI fields.
# For fields whose meaning is not documented we fall back to positional names.
# --------------------------------------------------------------------------- #
SCHEMAS = {
    "Harness circuit information": [
        "Circuit ID", "Rev", "Date", "Source", "Circuit ID (2)", "Rev (2)",
        "Date (2)", "F8", "F9", "F10", "Site", "F12", "F13", "F14", "F15",
        "F16", "F17", "F18", "F19", "F20", "F21", "F22", "F23", "Description",
        "F25", "F26", "F27", "Metric", "F29", "Flag A", "Flag B", "F32",
    ],  # trailing fields = option codes -> handled dynamically
    "Harness branch configuration": [
        "From Node", "From Ref", "From Coord", "To Node", "To Ref", "To Coord",
        "Length (mm)", "F8", "Diameter (mm)",
    ],
    "Harness wire specification": [
        "Wire Name", "Option Code", "Wire Spec", "Color", "Size (mm2)", "Class",
        "F7", "F8", "From Connector", "F10", "From Cavity", "From Term",
        "To Connector", "F14", "To Cavity", "To Term", "Flag A", "Flag B",
        "F19", "F20", "Twist/Dir", "F22", "F23", "F24", "Length Min (mm)",
        "Length Max (mm)", "F27", "F28", "Signal / Family", "F30", "Part Number",
    ],
    "Harness main node components": [
        "Reference", "F2", "F3", "F4", "Type", "F6", "Description", "F8",
        "Part Name", "Qty", "F11", "F12", "F13", "Part Number", "Flag A",
        "Flag B", "Flag C", "F18", "F19", "Process", "Family", "F22", "F23",
        "F24", "F25", "F26", "Config", "F28", "Kind", "Cav Count", "Flag D",
    ],
    "Harness branch insulations": [
        "From Node", "To Node", "Seq", "Diameter (mm)", "F5", "F6", "F7", "F8",
        "Source", "F10", "F11", "F12", "Order", "Type", "Material Spec",
        "Part Number", "F17", "Family", "F19", "Value", "F21", "F22", "F23",
        "F24", "F25", "F26", "F27", "F28", "F29", "F30", "Flag A", "Flag B",
    ],
    "Harness multicores": [
        "Multicore ID", "Flag A", "Flag B", "F4", "F5", "Dir 1", "Val 1",
        "Dir 2", "Val 2", "Dir 3", "Val 3", "F12", "F13", "Lay Length (mm)",
        "Twist", "F16", "F17", "F18", "F19", "Length Min (mm)", "Length Max (mm)",
    ],
    "Composite Option Codes": [
        "Option Code", "Description",
    ],
    "Harness wire through nodes": [
        "Wire Name", "Option Code", "Seq", "Node", "F5", "Through",
    ],
    "Harness branch insulation through nodes": [
        "Insulation ID", "Seq", "Node", "F4", "Through",
    ],
    "Harness Scope": [
        "Harness ID", "Rev", "Harness ID (2)", "Rev (2)", "Attribute", "Value",
    ],
}

# Nicer, shorter tab names (Excel limit = 31 chars, no : \ / ? * [ ])
SHEET_NAMES = {
    "Harness circuit information": "Circuit Information",
    "Harness branch configuration": "Branch Configuration",
    "Harness wire specification": "Wire Specification",
    "Harness main node components": "Main Node Components",
    "Harness branch insulations": "Branch Insulations",
    "Harness multicores": "Multicores",
    "Composite Option Codes": "Option Codes",
    "Harness wire through nodes": "Wire Through Nodes",
    "Harness branch insulation through nodes": "Insulation Through Nodes",
    "Harness Scope": "Harness Scope",
}


def parse_dsi(path):
    """Return (meta, sections) where sections is list of (name, headers, rows)."""
    with open(path, encoding="utf-8") as fh:
        raw = fh.read().splitlines()

    # --- file metadata from the header comment block ---
    meta = {}
    for ln in raw[:12]:
        m = re.match(r"!\s*\*?\s*([^:]+?):\s+(.*)", ln)
        if m and ":" in ln:
            key = m.group(1).strip()
            val = m.group(2).strip()
            if key and val and key not in ("Comment Marker", "Section Separator",
                                           "Field Delimiter"):
                meta[key] = val

    # --- split into sections ---
    idx = [i for i, l in enumerate(raw) if l.startswith("%")]
    sections = []
    raw_circuit = []
    for k, start in enumerate(idx):
        name = raw[start][1:].strip()
        end = idx[k + 1] if k + 1 < len(idx) else len(raw)
        data = [l for l in raw[start + 1:end]
                if l and not l.startswith("!")]
        if name == "Harness circuit information":
            raw_circuit = [l.split(DELIM) for l in data]
        if not data or name not in SCHEMAS:
            continue
        headers, rows = build_rows(name, data)
        sections.append((name, headers, rows))
    return meta, sections, raw_circuit


def build_rows(name, data):
    base = SCHEMAS[name]
    split = [r.split(DELIM) for r in data]
    maxlen = max(len(r) for r in split)

    if name == "Harness circuit information":
        # fields after the base schema are a variable list of option codes
        extra = maxlen - len(base)
        headers = base + [f"Option {i+1}" for i in range(max(extra, 0))]
    else:
        headers = list(base)
        # pad header list if some rows carry more columns than documented
        while len(headers) < maxlen:
            headers.append(f"F{len(headers)+1}")

    rows = []
    for r in split:
        r = r + [""] * (len(headers) - len(r))
        rows.append([c.strip() for c in r[:len(headers)]])

    # drop columns that are completely empty across every row (keep table tidy)
    keep = [i for i in range(len(headers))
            if headers[i] not in ("",) and any(row[i] for row in rows)]
    # always keep the first column even if it looked empty
    if 0 not in keep:
        keep.insert(0, 0)
    headers = [headers[i] for i in keep]
    rows = [[row[i] for i in keep] for row in rows]
    return headers, rows


# --------------------------------------------------------------------------- #
# Minimal XLSX writer (SpreadsheetML)
# --------------------------------------------------------------------------- #
def col_letter(n):  # 1 -> A, 27 -> AA
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


NUMERIC_RE = re.compile(r"^-?\d+(\.\d+)?$")


def is_number(v):
    return bool(NUMERIC_RE.match(v)) and v not in ("", "-")


def cell_xml(ref, value, style, force_text=False):
    if value == "":
        return f'<c r="{ref}" s="{style}"/>'
    if not force_text and is_number(value):
        return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'
    return (f'<c r="{ref}" s="{style}" t="inlineStr">'
            f'<is><t xml:space="preserve">{escape(value)}</t></is></c>')


def sheet_xml(headers, rows, force_text_cols, centered_cols=None, narrow_cols=None):
    centered_cols = centered_cols or set()
    narrow_cols = narrow_cols or set()
    ncols = len(headers)
    last_col = col_letter(ncols)
    nrows = len(rows) + 1
    dim = f"A1:{last_col}{nrows}"

    # column widths from content
    widths = [len(str(h)) for h in headers]
    for row in rows:
        for i, v in enumerate(row):
            if len(v) > widths[i]:
                widths[i] = len(v)
    cols = ['<cols>']
    for i, w in enumerate(widths, 1):
        if (i - 1) in narrow_cols:
            width = max(len(str(headers[i - 1])) + 2, 7)
        else:
            width = min(max(w + 2, 9), 60)
        cols.append(f'<col min="{i}" max="{i}" width="{width}" customWidth="1"/>')
    cols.append('</cols>')

    out = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        f'<dimension ref="{dim}"/>',
        '<sheetViews><sheetView workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '</sheetView></sheetViews>',
        '<sheetFormatPr defaultRowHeight="15"/>',
        "".join(cols),
        '<sheetData>',
    ]
    # header row (style 2 = header)
    hdr = [f'<row r="1">']
    for i, h in enumerate(headers, 1):
        hdr.append(cell_xml(f"{col_letter(i)}1", str(h), 2, force_text=True))
    hdr.append('</row>')
    out.append("".join(hdr))
    # data rows (style 1 = normal border, style 3 = centered)
    for ri, row in enumerate(rows, 2):
        cells = [f'<row r="{ri}">']
        for ci, v in enumerate(row, 1):
            col0 = ci - 1
            ft = col0 in force_text_cols
            style = 3 if col0 in centered_cols else 1
            cells.append(cell_xml(f"{col_letter(ci)}{ri}", v, style, force_text=ft))
        cells.append('</row>')
        out.append("".join(cells))
    out.append('</sheetData>')
    out.append(f'<autoFilter ref="{dim}"/>')
    out.append('<tableParts count="1"><tablePart r:id="rId1"/></tableParts>')
    out.append('</worksheet>')
    return "".join(out)


def table_xml(tid, name, headers, nrows):
    last_col = col_letter(len(headers))
    ref = f"A1:{last_col}{nrows + 1}"
    cols = "".join(
        f'<tableColumn id="{i}" name="{escape(uniquify(h, headers, i-1))}"/>'
        for i, h in enumerate(headers, 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'id="{tid}" name="{name}" displayName="{name}" ref="{ref}" '
        f'totalsRowShown="0">'
        f'<autoFilter ref="{ref}"/>'
        f'<tableColumns count="{len(headers)}">{cols}</tableColumns>'
        '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" '
        'showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>'
        '</table>'
    )


def uniquify(name, headers, idx):
    """Excel table columns must have unique, non-empty names."""
    name = name or f"Col{idx+1}"
    seen = headers[:idx]
    if seen.count(name):
        return f"{name} {seen.count(name)+1}"
    return name


def styles_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2">'
        '<font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
        '</fonts>'
        '<fills count="3">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF305496"/>'
        '<bgColor indexed="64"/></patternFill></fill>'
        '</fills>'
        '<borders count="2">'
        '<border><left/><right/><top/><bottom/><diagonal/></border>'
        '<border><left style="thin"><color rgb="FFD9D9D9"/></left>'
        '<right style="thin"><color rgb="FFD9D9D9"/></right>'
        '<top style="thin"><color rgb="FFD9D9D9"/></top>'
        '<bottom style="thin"><color rgb="FFD9D9D9"/></bottom><diagonal/></border>'
        '</borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="4">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" '
        'applyFont="1" applyFill="1" applyBorder="1">'
        '<alignment horizontal="center" vertical="center"/></xf>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" '
        'applyBorder="1" applyAlignment="1">'
        '<alignment horizontal="center"/></xf>'
        '</cellXfs>'
        '</styleSheet>'
    )


# In the raw "Harness circuit information" record, option codes start at this
# field index; field 26 holds the human-readable description.
CIRCUIT_DESC_FIELD = 26
CIRCUIT_OPTION_START = 36
_OPTION_RE = re.compile(r"^[A-Z]\d[A-Z0-9]{3}$|^[A-Z0-9]{5}$")


def build_option_matrix(raw_records):
    """Build the Ref x Option-code presence matrix from raw circuit records.

    Rows   = each circuit (Reference + Description)
    Cols   = every unique option code (de-duplicated, sorted)
    Cell   = 'X' if the code is in that circuit's option list, else empty.

    We read straight from the raw DSI fields so only genuine option codes are
    used (skipping flags like 'true'/'false' and empty positions).
    """
    if not raw_records:
        return None

    circuits = []
    all_codes = []
    seen = set()
    for f in raw_records:
        ref = f[0].strip() if f else ""
        desc = f[CIRCUIT_DESC_FIELD].strip() if len(f) > CIRCUIT_DESC_FIELD else ""
        codes = set()
        for c in f[CIRCUIT_OPTION_START:]:
            c = c.strip()
            if not c or c.lower() in ("true", "false"):
                continue
            codes.add(c)
            if c not in seen:
                seen.add(c)
                all_codes.append(c)
        circuits.append((ref, desc, codes))
    all_codes.sort()

    hdr = ["Reference", "Description"] + all_codes
    out_rows = []
    for ref, desc, codes in circuits:
        row = [ref, desc] + ["X" if c in codes else "" for c in all_codes]
        out_rows.append(row)
    return hdr, out_rows


def build_workbook(path, meta, sections, out_path, raw_circuit=None):
    # Build an "Overview" sheet as the first sheet
    overview_rows = [[k, v] for k, v in meta.items()]
    overview_rows.append(["", ""])
    overview_rows.append(["Section", "Rows"])
    for name, headers, rows in sections:
        overview_rows.append([SHEET_NAMES.get(name, name), str(len(rows))])
    sheets = [("Overview", ["Property", "Value"], overview_rows, set())]

    # columns to center/narrow, keyed by sheet display name
    center_map = {}
    narrow_map = {}

    # Option-code presence matrix (Ref x Codes, X where present)
    matrix = build_option_matrix(raw_circuit)
    if matrix:
        mh, mr = matrix
        # keep Reference as text; X columns are plain text too
        sheets.append(("Option Matrix", mh, mr, set(range(len(mh)))))
        # center + narrow every option-code column (index 2 onward)
        center_map["Option Matrix"] = set(range(2, len(mh)))
        narrow_map["Option Matrix"] = set(range(2, len(mh)))

    for name, headers, rows in sections:
        # force-text columns: identifiers/codes that look numeric but must stay text
        ftc = set()
        for i, h in enumerate(headers):
            hl = h.lower()
            if any(t in hl for t in ("id", "code", "cavity", "rev", "part",
                                     "reference", "node", "name")):
                ftc.add(i)
        sheets.append((SHEET_NAMES.get(name, name), headers, rows, ftc))

    # de-duplicate / trim sheet names to <=31 chars
    used = {}
    final = []
    for sn, headers, rows, ftc in sheets:
        base = re.sub(r'[:\\/?*\[\]]', ' ', sn)[:31]
        nm = base
        n = 2
        while nm in used:
            nm = f"{base[:28]} {n}"
            n += 1
        used[nm] = True
        final.append((nm, headers, rows, ftc))
    sheets = final

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        # [Content_Types].xml
        overrides = [
            '<Override PartName="/xl/workbook.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
            '<Override PartName="/xl/styles.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
        ]
        for i in range(1, len(sheets) + 1):
            overrides.append(
                f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
                'ContentType="application/vnd.openxmlformats-officedocument.'
                'spreadsheetml.worksheet+xml"/>')
            overrides.append(
                f'<Override PartName="/xl/tables/table{i}.xml" '
                'ContentType="application/vnd.openxmlformats-officedocument.'
                'spreadsheetml.table+xml"/>')
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
                   'content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.'
                   'openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   + "".join(overrides) + '</Types>')

        # _rels/.rels
        z.writestr("_rels/.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/'
                   'package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.'
                   'org/officeDocument/2006/relationships/officeDocument" '
                   'Target="xl/workbook.xml"/></Relationships>')

        # xl/workbook.xml
        sheet_tags = "".join(
            f'<sheet name="{escape(nm)}" sheetId="{i}" r:id="rId{i}"/>'
            for i, (nm, *_ ) in enumerate(sheets, 1))
        z.writestr("xl/workbook.xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<workbook xmlns="http://schemas.openxmlformats.org/'
                   'spreadsheetml/2006/main" xmlns:r="http://schemas.'
                   'openxmlformats.org/officeDocument/2006/relationships">'
                   f'<sheets>{sheet_tags}</sheets></workbook>')

        # xl/_rels/workbook.xml.rels
        rels = []
        for i in range(1, len(sheets) + 1):
            rels.append(f'<Relationship Id="rId{i}" Type="http://schemas.'
                        'openxmlformats.org/officeDocument/2006/relationships/'
                        f'worksheet" Target="worksheets/sheet{i}.xml"/>')
        sid = len(sheets) + 1
        rels.append(f'<Relationship Id="rId{sid}" Type="http://schemas.'
                    'openxmlformats.org/officeDocument/2006/relationships/'
                    'styles" Target="styles.xml"/>')
        z.writestr("xl/_rels/workbook.xml.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/'
                   'package/2006/relationships">' + "".join(rels) +
                   '</Relationships>')

        z.writestr("xl/styles.xml", styles_xml())

        # sheets + tables
        for i, (nm, headers, rows, ftc) in enumerate(sheets, 1):
            z.writestr(f"xl/worksheets/sheet{i}.xml",
                       sheet_xml(headers, rows, ftc,
                                 centered_cols=center_map.get(nm),
                                 narrow_cols=narrow_map.get(nm)))
            z.writestr(f"xl/worksheets/_rels/sheet{i}.xml.rels",
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<Relationships xmlns="http://schemas.openxmlformats.org/'
                       'package/2006/relationships">'
                       f'<Relationship Id="rId1" Type="http://schemas.'
                       'openxmlformats.org/officeDocument/2006/relationships/'
                       f'table" Target="../tables/table{i}.xml"/>'
                       '</Relationships>')
            tname = "T_" + re.sub(r'[^A-Za-z0-9_]', '_', nm)
            z.writestr(f"xl/tables/table{i}.xml",
                       table_xml(i, tname, headers, len(rows)))

    return sheets


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else "HQRNESS2 TEST.dsi"
    out = sys.argv[2] if len(sys.argv) > 2 else "HARNESS2_TEST.xlsx"
    meta, sections, raw_circuit = parse_dsi(src)
    sheets = build_workbook(src, meta, sections, out, raw_circuit)
    print(f"Wrote {out}")
    print(f"Metadata fields: {len(meta)}")
    for nm, headers, rows, ftc in sheets:
        print(f"  {nm:28s} cols={len(headers):3d}  rows={len(rows)}")


if __name__ == "__main__":
    main()
