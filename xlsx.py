#!/usr/bin/env python3
"""Minimal .xlsx writer with per-cell fills. Standard library only.

An .xlsx is a zip of XML parts, so it can be built directly. This exists
alongside ``dsi_to_excel.py`` rather than replacing it because the comparison
report needs something that module cannot do: colour an individual cell to mark
a difference. ``dsi_to_excel.py`` is left untouched so it keeps reproducing its
committed workbook byte for byte.

Usage::

    sheet = Sheet("Wires", ["Wire", "Colour"])
    sheet.add(["1062-IM", "BA"])
    sheet.add(["1063-IM", "NR"], fills={1: YELLOW})
    write_workbook("out.xlsx", [sheet])

Style names are the module constants below. Anything else raises, so a typo
fails loudly instead of silently producing an unstyled cell.
"""

from __future__ import annotations

import re
import zipfile
from html import escape
from typing import Dict, Iterable, List, Optional, Sequence, Set

__all__ = [
    "Sheet",
    "write_workbook",
    "PLAIN",
    "YELLOW",
    "RED",
    "GREEN",
    "GREY",
    "BOLD",
    "CENTER",
]

# Style keys, mapped to cellXfs indices in _CELL_XFS order.
PLAIN = "plain"
CENTER = "center"
YELLOW = "yellow"  # modified
RED = "red"        # removed
GREEN = "green"    # added
GREY = "grey"      # not compared
BOLD = "bold"

# cellXfs index per style. Index 0 is bare, 1 is the bordered body style,
# 2 the header. Excel's own Bad/Good/Neutral colours are used rather than the
# VBA's pure 255 / 65535 / 65280, which are unreadable behind black text.
_STYLE_INDEX = {
    PLAIN: 1,
    "header": 2,
    CENTER: 3,
    YELLOW: 4,
    RED: 5,
    GREEN: 6,
    GREY: 7,
    BOLD: 8,
}

_NUMERIC = re.compile(r"^-?\d+(\.\d+)?$")
# XML 1.0 forbids most control characters outright.
_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_RELS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT = "http://schemas.openxmlformats.org/package/2006/content-types"


class Sheet:
    """One worksheet: a header row plus body rows, with optional cell fills.

    Args:
        name: tab name. Sanitised and truncated to Excel's 31 characters.
        headers: column titles. Duplicates are made unique for the table part.
        text_cols: column indices to force to text. Part numbers and cavity
            numbers look numeric but must not be right-aligned or stripped of
            leading zeros.
        center_cols: column indices to centre.
        table: wrap in a native Excel Table (banded rows plus filter).
            Switched off for layouts with merged concerns, like a synthesis
            sheet with blank spacer rows.
        widths: explicit column widths, overriding the fitted ones.
    """

    def __init__(
        self,
        name: str,
        headers: Sequence[str],
        text_cols: Optional[Iterable[int]] = None,
        center_cols: Optional[Iterable[int]] = None,
        table: bool = True,
        widths: Optional[Dict[int, float]] = None,
    ) -> None:
        self.name = name
        self.headers = list(headers)
        self.rows: List[List[str]] = []
        self.fills: List[Dict[int, str]] = []
        self.text_cols: Set[int] = set(text_cols or ())
        self.center_cols: Set[int] = set(center_cols or ())
        self.table = table
        self.widths = dict(widths or {})

    def add(self, row: Sequence[object], fills: Optional[Dict[int, str]] = None) -> None:
        """Append a row. ``fills`` maps column index to a style constant."""
        values = ["" if v is None else str(v) for v in row]
        if len(values) < len(self.headers):
            values += [""] * (len(self.headers) - len(values))
        self.rows.append(values[: len(self.headers)])
        for style in (fills or {}).values():
            if style not in _STYLE_INDEX:
                raise ValueError("unknown cell style {!r}".format(style))
        self.fills.append(dict(fills or {}))

    def __len__(self) -> int:
        return len(self.rows)


def _col_letter(n: int) -> str:
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def _clean(value: str) -> str:
    return _ILLEGAL.sub("", value)


def _cell(ref: str, value: str, style: int, as_text: bool) -> str:
    if value == "":
        return '<c r="{}" s="{}"/>'.format(ref, style)
    if not as_text and _NUMERIC.match(value):
        return '<c r="{}" s="{}"><v>{}</v></c>'.format(ref, style, value)
    return (
        '<c r="{}" s="{}" t="inlineStr"><is><t xml:space="preserve">{}</t></is></c>'
    ).format(ref, style, escape(_clean(value)))


def _unique_headers(headers: Sequence[str]) -> List[str]:
    """Table column names must be unique and non-empty."""
    out: List[str] = []
    for index, name in enumerate(headers):
        name = _clean(str(name)).strip() or "Col{}".format(index + 1)
        candidate = name
        suffix = 2
        while candidate in out:
            candidate = "{} {}".format(name, suffix)
            suffix += 1
        out.append(candidate)
    return out


def _sheet_xml(sheet: Sheet) -> str:
    ncols = max(len(sheet.headers), 1)
    nrows = len(sheet.rows) + 1
    dim = "A1:{}{}".format(_col_letter(ncols), nrows)

    widths = [len(str(h)) for h in sheet.headers]
    for row in sheet.rows:
        for i, value in enumerate(row):
            if i < len(widths) and len(value) > widths[i]:
                widths[i] = len(value)
    cols = ["<cols>"]
    for i, width in enumerate(widths, 1):
        chosen = sheet.widths.get(i - 1, min(max(width + 2, 9), 55))
        cols.append('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>'.format(i, chosen))
    cols.append("</cols>")

    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="{}" xmlns:r="{}">'.format(_MAIN, _RELS),
        '<dimension ref="{}"/>'.format(dim),
        '<sheetViews><sheetView workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        "</sheetView></sheetViews>",
        '<sheetFormatPr defaultRowHeight="15"/>',
        "".join(cols) if sheet.headers else "",
        "<sheetData>",
    ]

    header_cells = ['<row r="1">']
    for i, title in enumerate(sheet.headers, 1):
        header_cells.append(_cell("{}1".format(_col_letter(i)), str(title), _STYLE_INDEX["header"], True))
    header_cells.append("</row>")
    parts.append("".join(header_cells))

    for r, (row, fills) in enumerate(zip(sheet.rows, sheet.fills), 2):
        cells = ['<row r="{}">'.format(r)]
        for c, value in enumerate(row, 1):
            col0 = c - 1
            style_name = fills.get(col0)
            if style_name:
                style = _STYLE_INDEX[style_name]
            elif col0 in sheet.center_cols:
                style = _STYLE_INDEX[CENTER]
            else:
                style = _STYLE_INDEX[PLAIN]
            cells.append(_cell("{}{}".format(_col_letter(c), r), value, style, col0 in sheet.text_cols))
        cells.append("</row>")
        parts.append("".join(cells))

    parts.append("</sheetData>")
    use_table = sheet.table and sheet.headers and sheet.rows
    if not use_table and sheet.headers:
        # No table, so the filter has to live on the sheet itself.
        parts.append('<autoFilter ref="{}"/>'.format(dim))
    if use_table:
        parts.append('<tableParts count="1"><tablePart r:id="rId1"/></tableParts>')
    parts.append("</worksheet>")
    return "".join(parts)


def _table_xml(table_id: int, display_name: str, headers: Sequence[str], nrows: int) -> str:
    ref = "A1:{}{}".format(_col_letter(len(headers)), nrows + 1)
    columns = "".join(
        '<tableColumn id="{}" name="{}"/>'.format(i, escape(name))
        for i, name in enumerate(_unique_headers(headers), 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<table xmlns="{main}" id="{tid}" name="{dn}" displayName="{dn}" ref="{ref}" '
        'totalsRowShown="0"><autoFilter ref="{ref}"/>'
        '<tableColumns count="{n}">{cols}</tableColumns>'
        '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" '
        'showRowStripes="1" showColumnStripes="0"/></table>'
    ).format(main=_MAIN, tid=table_id, dn=display_name, ref=ref, n=len(headers), cols=columns)


def _styles_xml() -> str:
    fills = [
        '<fill><patternFill patternType="none"/></fill>',
        '<fill><patternFill patternType="gray125"/></fill>',
        _solid("FF305496"),  # header
        _solid("FFFFEB9C"),  # yellow / modified
        _solid("FFFFC7CE"),  # red / removed
        _solid("FFC6EFCE"),  # green / added
        _solid("FFF2F2F2"),  # grey / not compared
    ]
    fonts = [
        '<font><sz val="11"/><name val="Calibri"/></font>',
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>',
        '<font><sz val="11"/><color rgb="FF9C6500"/><name val="Calibri"/></font>',
        '<font><sz val="11"/><color rgb="FF9C0006"/><name val="Calibri"/></font>',
        '<font><sz val="11"/><color rgb="FF006100"/><name val="Calibri"/></font>',
        '<font><b/><sz val="11"/><name val="Calibri"/></font>',
    ]
    # Order must match _STYLE_INDEX.
    xfs = [
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>',
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" '
        'applyFill="1" applyBorder="1"><alignment horizontal="center" vertical="center" '
        'wrapText="1"/></xf>',
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" '
        'applyAlignment="1"><alignment horizontal="center"/></xf>',
        '<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" '
        'applyFill="1" applyBorder="1"/>',
        '<xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" '
        'applyFill="1" applyBorder="1"/>',
        '<xf numFmtId="0" fontId="4" fillId="5" borderId="1" xfId="0" applyFont="1" '
        'applyFill="1" applyBorder="1"/>',
        '<xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0" applyFill="1" '
        'applyBorder="1"/>',
        '<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
    ]
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="{main}">'
        '<fonts count="{nf}">{fonts}</fonts>'
        '<fills count="{nfl}">{fills}</fills>'
        '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
        '<border><left style="thin"><color rgb="FFD9D9D9"/></left>'
        '<right style="thin"><color rgb="FFD9D9D9"/></right>'
        '<top style="thin"><color rgb="FFD9D9D9"/></top>'
        '<bottom style="thin"><color rgb="FFD9D9D9"/></bottom><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="{nx}">{xfs}</cellXfs>'
        "</styleSheet>"
    ).format(
        main=_MAIN,
        nf=len(fonts),
        fonts="".join(fonts),
        nfl=len(fills),
        fills="".join(fills),
        nx=len(xfs),
        xfs="".join(xfs),
    )


def _solid(rgb: str) -> str:
    return (
        '<fill><patternFill patternType="solid"><fgColor rgb="{}"/>'
        '<bgColor indexed="64"/></patternFill></fill>'
    ).format(rgb)


def _safe_names(sheets: Sequence[Sheet]) -> List[str]:
    used: Set[str] = set()
    names = []
    for sheet in sheets:
        base = re.sub(r"[:\\/?*\[\]]", " ", sheet.name).strip()[:31] or "Sheet"
        name = base
        suffix = 2
        while name in used:
            name = "{} {}".format(base[:28], suffix)
            suffix += 1
        used.add(name)
        names.append(name)
    return names


def write_workbook(path: str, sheets: Sequence[Sheet]) -> None:
    """Write ``sheets`` to ``path``. Raises ValueError when there are none."""
    if not sheets:
        raise ValueError("a workbook needs at least one sheet")
    names = _safe_names(sheets)

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        overrides = [
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.'
            'openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.'
            'openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
        ]
        for i, sheet in enumerate(sheets, 1):
            overrides.append(
                '<Override PartName="/xl/worksheets/sheet{}.xml" ContentType="application/vnd.'
                'openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.format(i)
            )
            if sheet.table and sheet.headers and sheet.rows:
                overrides.append(
                    '<Override PartName="/xl/tables/table{}.xml" ContentType="application/vnd.'
                    'openxmlformats-officedocument.spreadsheetml.table+xml"/>'.format(i)
                )
        z.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="{ct}">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.'
            'relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
            "{ov}</Types>".format(ct=_CT, ov="".join(overrides)),
        )

        z.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="{pkg}"><Relationship Id="rId1" Type="{rels}/officeDocument" '
            'Target="xl/workbook.xml"/></Relationships>'.format(pkg=_PKG_RELS, rels=_RELS),
        )

        tags = "".join(
            '<sheet name="{}" sheetId="{}" r:id="rId{}"/>'.format(escape(name), i, i)
            for i, name in enumerate(names, 1)
        )
        z.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="{main}" xmlns:r="{rels}"><sheets>{tags}</sheets></workbook>'.format(
                main=_MAIN, rels=_RELS, tags=tags
            ),
        )

        rels = [
            '<Relationship Id="rId{0}" Type="{1}/worksheet" Target="worksheets/sheet{0}.xml"/>'.format(i, _RELS)
            for i in range(1, len(sheets) + 1)
        ]
        rels.append(
            '<Relationship Id="rId{}" Type="{}/styles" Target="styles.xml"/>'.format(len(sheets) + 1, _RELS)
        )
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="{pkg}">{rels}</Relationships>'.format(pkg=_PKG_RELS, rels="".join(rels)),
        )

        z.writestr("xl/styles.xml", _styles_xml())

        for i, (name, sheet) in enumerate(zip(names, sheets), 1):
            z.writestr("xl/worksheets/sheet{}.xml".format(i), _sheet_xml(sheet))
            if sheet.table and sheet.headers and sheet.rows:
                z.writestr(
                    "xl/worksheets/_rels/sheet{}.xml.rels".format(i),
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<Relationships xmlns="{pkg}"><Relationship Id="rId1" Type="{rels}/table" '
                    'Target="../tables/table{i}.xml"/></Relationships>'.format(pkg=_PKG_RELS, rels=_RELS, i=i),
                )
                display = "T_" + re.sub(r"[^A-Za-z0-9_]", "_", name)
                z.writestr(
                    "xl/tables/table{}.xml".format(i),
                    _table_xml(i, display, sheet.headers, len(sheet.rows)),
                )
