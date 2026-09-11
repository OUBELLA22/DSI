#!/usr/bin/env python3
"""Element views for single-file mode: named, filtered projections of a DSI.

Each :class:`View` says which section to read, which rows to keep, and which
fields to show under which heading. Field indices were established by reading
the sample export rather than trusting the undocumented layout -- notably:

* main node components field 4 is the type discriminator
  (``CONNECTOR`` / ``SPLICE`` / ``IDC``), which is what separates connectors
  from splices;
* circuit option codes start at field 36 (all 1704 tokens in the sample resolve
  to entries in ``Composite Option Codes``);
* branch coordinates are 3D, packed as ``x1820.00y-905.00z0.00``.

That last one matters. The VBA's ``grapheCAPH`` splits the coordinate on ``"y"``
only, so Y comes out as ``-905.00z0.00``, ``CSng`` throws, and
``On Error Resume Next`` leaves the column silently blank. Here the coordinate
is parsed with an explicit pattern and anything unparseable is reported.
"""

from __future__ import annotations

import re
from typing import Callable, Dict, List, Optional, Sequence, Tuple

from dsi import Dsi, Section

# x/y/z packed into one field. Signs are common, so the digits pattern must
# allow a leading minus rather than splitting on the letters.
_COORD = re.compile(r"x(-?[\d.]+)y(-?[\d.]+)z(-?[\d.]+)", re.IGNORECASE)

# Headings whose values must stay text: they look numeric but leading zeros and
# exact digits matter.
_TEXT_HINTS = ("id", "code", "cavity", "rev", "part", "reference", "node", "name", "circuit", "wire")


class Ctx:
    """Lookup helper shared by derived columns, so counts are computed once."""

    def __init__(self, dsi: Dsi) -> None:
        self.dsi = dsi
        self._counts: Dict[Tuple[str, int], Dict[str, int]] = {}

    def counts(self, section_name: str, field: int) -> Dict[str, int]:
        cache_key = (section_name, field)
        if cache_key not in self._counts:
            tally: Dict[str, int] = {}
            section = self.dsi.section(section_name)
            if section is not None:
                for row in section.rows:
                    value = section.field(row, field)
                    if value:
                        tally[value] = tally.get(value, 0) + 1
            self._counts[cache_key] = tally
        return self._counts[cache_key]

    def wire_ends(self) -> Dict[str, int]:
        """How many wires land on each node reference (either end)."""
        cache_key = ("__wire_ends__", 0)
        if cache_key not in self._counts:
            tally: Dict[str, int] = {}
            wires = self.dsi.section("Harness wire specification")
            if wires is not None:
                for row in wires.rows:
                    for field in (8, 12):
                        ref = wires.field(row, field)
                        if ref:
                            tally[ref] = tally.get(ref, 0) + 1
            self._counts[cache_key] = tally
        return self._counts[cache_key]


# A column is a heading plus either a field index or a function of (row, ctx).
Spec = object  # int | Callable[[Section, Sequence[str], Ctx], str]


class View:
    """One menu entry: a named table built from one section."""

    def __init__(
        self,
        key: str,
        label: str,
        section: Sequence[str],
        columns: Sequence[Tuple[str, Spec]],
        group: str = "Other",
        sheet: Optional[str] = None,
        where: Optional[Callable[[Section, Sequence[str]], bool]] = None,
    ) -> None:
        self.key = key
        self.label = label
        self.section_names = tuple(section)
        self.columns = list(columns)
        self.group = group
        self.sheet = sheet or label
        self.where = where

    def resolve(self, dsi: Dsi) -> Optional[Section]:
        return dsi.find(*self.section_names)

    def count(self, dsi: Dsi) -> int:
        section = self.resolve(dsi)
        if section is None:
            return 0
        if self.where is None:
            return len(section.rows)
        return sum(1 for row in section.rows if self.where(section, row))

    def build(self, dsi: Dsi, ctx: Ctx) -> Tuple[List[str], List[List[str]]]:
        """Return ``(headers, rows)`` for this view, empty when absent."""
        section = self.resolve(dsi)
        if section is None:
            return [c[0] for c in self.columns], []
        headers = [c[0] for c in self.columns]
        out: List[List[str]] = []
        for row in section.rows:
            if self.where is not None and not self.where(section, row):
                continue
            built = []
            for _, spec in self.columns:
                if isinstance(spec, int):
                    built.append(section.field(row, spec))
                else:
                    built.append(spec(section, row, ctx))
            out.append(built)
        return headers, out

    def text_columns(self) -> set:
        return {
            i
            for i, (title, _) in enumerate(self.columns)
            if any(hint in title.lower() for hint in _TEXT_HINTS)
        }


# -- derived column helpers -------------------------------------------------


def _coord(field: int, part: int) -> Callable:
    """Extract x, y or z (``part`` 0/1/2) from a packed coordinate field."""

    def get(section: Section, row: Sequence[str], ctx: Ctx) -> str:
        raw = section.field(row, field)
        if not raw:
            return ""
        match = _COORD.search(raw)
        if not match:
            # Loud rather than blank: an unparsed coordinate is a real problem.
            return "?? " + raw
        return match.group(part + 1)

    return get


def _count_from(section_name: str, field: int) -> Callable:
    def get(section: Section, row: Sequence[str], ctx: Ctx) -> str:
        ref = section.field(row, 0)
        n = ctx.counts(section_name, field).get(ref, 0)
        return str(n) if n else ""

    return get


def _wire_count(section: Section, row: Sequence[str], ctx: Ctx) -> str:
    n = ctx.wire_ends().get(section.field(row, 0), 0)
    return str(n) if n else ""


def _option_codes(start: int) -> Callable:
    """Join the trailing variable-length option code list into one cell."""

    def get(section: Section, row: Sequence[str], ctx: Ctx) -> str:
        codes = [v.strip() for v in row[start:] if v.strip()]
        return " ".join(codes)

    return get


def _by_type(*wanted: str) -> Callable:
    allowed = {w.upper() for w in wanted}

    def keep(section: Section, row: Sequence[str]) -> bool:
        return section.field(row, 4).upper() in allowed

    return keep


# -- node component columns -------------------------------------------------

_NODE_COMMON: List[Tuple[str, Spec]] = [
    ("Reference", 0),
    ("Type", 4),
    ("Description", 6),
    ("Part Name", 8),
    ("Part Number", 12),
    ("Family", 19),
    ("Colour", 27),
    ("Cavities", 28),
    ("Terminals", _count_from("terminals", 0)),
    ("Seals", _count_from("seals", 0)),
    ("Plugs", _count_from("plug", 0)),
    ("Wires", _wire_count),
    ("Route", 7),
]

_ACCESSORY: List[Tuple[str, Spec]] = [
    ("Reference", 0),
    ("Type", 4),
    ("Option Expression", 5),
    ("Parent / Node", 6),
    ("Part Name", 8),
    ("Qty", 9),
    ("Part Number", 12),
    ("Family", 19),
    ("Colour", 27),
]

VIEWS: List[View] = [
    # -- components ---------------------------------------------------------
    View(
        "connectors",
        "Connectors",
        ["Harness main node components"],
        _NODE_COMMON,
        group="Components",
        where=_by_type("CONNECTOR"),
    ),
    View(
        "splices",
        "Splices",
        ["Harness main node components"],
        _NODE_COMMON,
        group="Components",
        where=_by_type("SPLICE"),
    ),
    View(
        "idc",
        "IDC connections",
        ["Harness main node components"],
        _NODE_COMMON,
        group="Components",
        where=_by_type("IDC"),
    ),
    View(
        "nodes",
        "All node components",
        ["Harness main node components"],
        _NODE_COMMON,
        group="Components",
    ),
    View(
        "terminals",
        "Terminals",
        ["terminals", "Harness terminals"],
        [
            ("Connector Reference", 0),
            ("Cavity", 1),
            ("Type", 4),
            ("Part Name", 8),
            ("Qty", 9),
            ("Part Number", 12),
            ("Sealed", 13),
            ("Selection", 17),
            ("Plating", 18),
        ],
        group="Components",
    ),
    View(
        "seals",
        "Cavity seals",
        ["seals", "Harness cavity seals"],
        [
            ("Connector Reference", 0),
            ("Cavity", 1),
            ("Type", 4),
            ("Part Name", 8),
            ("Qty", 9),
            ("Part Number", 12),
        ],
        group="Components",
    ),
    View(
        "plugs",
        "Cavity plugs",
        ["plug", "Harness cavity plugs"],
        [
            ("Connector Reference", 0),
            ("Cavity", 1),
            ("Type", 4),
            ("Part Name", 8),
            ("Qty", 9),
            ("Part Number", 12),
        ],
        group="Components",
    ),
    View("clips", "Clips / fasteners", ["clips"], _ACCESSORY, group="Components"),
    View("grommets", "Grommets", ["grommets"], _ACCESSORY, group="Components"),
    View(
        "other",
        "Other components (tape, labels)",
        ["other components"],
        _ACCESSORY,
        group="Components",
        sheet="Other Components",
    ),
    View(
        "extra",
        "Extra node components",
        ["Harness extra node components"],
        _ACCESSORY,
        group="Components",
    ),
    # -- wiring -------------------------------------------------------------
    View(
        "wires",
        "Wires",
        ["Harness wire specification"],
        [
            ("Wire Name", 0),
            ("Option Expression", 1),
            ("Wire Spec", 2),
            ("Colour", 3),
            ("Size mm2", 4),
            ("Class", 5),
            ("Multicore", 7),
            ("From Connector", 8),
            ("From Cavity", 10),
            ("From Plating", 11),
            ("To Connector", 12),
            ("To Cavity", 14),
            ("To Plating", 15),
            ("Length Min", 24),
            ("Length Max", 25),
            ("Family", 28),
            ("Part Number", 30),
        ],
        group="Wiring",
    ),
    View(
        "multicores",
        "Multicores / twisted",
        ["Harness multicores"],
        [
            ("Multicore ID", 0),
            ("Flag A", 1),
            ("Flag B", 2),
            ("Lay Length", 13),
            ("Twisted", 14),
            ("Length Min", 20),
            ("Length Max", 21),
            ("Group", 22),
            ("Part Number", 24),
        ],
        group="Wiring",
    ),
    View(
        "wirenodes",
        "Wire through nodes",
        ["Harness wire through nodes"],
        [
            ("Wire Name", 0),
            ("Option Expression", 1),
            ("Sequence", 2),
            ("Node", 3),
            ("Through", 5),
        ],
        group="Wiring",
    ),
    # -- structure ----------------------------------------------------------
    View(
        "branches",
        "Branches (segments)",
        ["Harness branch configuration"],
        [
            ("From Node", 0),
            ("From X", _coord(2, 0)),
            ("From Y", _coord(2, 1)),
            ("From Z", _coord(2, 2)),
            ("To Node", 3),
            ("To X", _coord(5, 0)),
            ("To Y", _coord(5, 1)),
            ("To Z", _coord(5, 2)),
            ("Length mm", 6),
            ("Option Expression", 7),
            ("Diameter mm", 8),
        ],
        group="Structure",
    ),
    View(
        "insulations",
        "Insulations (tape / tube)",
        ["Harness branch insulations"],
        [
            ("From Node", 0),
            ("To Node", 2),
            ("Route", 3),
            ("Sequence", 4),
            ("Diameter mm", 6),
            ("Option Expression", 7),
            ("Source", 12),
            ("Order", 17),
            ("Type", 18),
            ("Material Spec", 19),
            ("Part Number", 21),
            ("Colour", 22),
            ("Family", 23),
            ("Coverage", 25),
        ],
        group="Structure",
    ),
    View(
        "insulnodes",
        "Insulation through nodes",
        ["Harness branch insulation through nodes"],
        [
            ("Insulation From", 0),
            ("Sequence", 1),
            ("Node", 2),
            ("Through", 4),
        ],
        group="Structure",
    ),
    # -- variance -----------------------------------------------------------
    View(
        "circuits",
        "Circuits / diversity",
        ["Harness circuit information"],
        [
            ("Circuit ID", 0),
            ("Rev", 1),
            ("Date", 2),
            ("Source", 3),
            ("Site", 11),
            ("Description", 26),
            ("Metric", 31),
            ("Option Codes", _option_codes(36)),
        ],
        group="Variance",
    ),
    View(
        "options",
        "Option code definitions",
        ["Composite Option Codes"],
        [("Option Code", 0), ("Description", 1)],
        group="Variance",
        sheet="Option Codes",
    ),
    View(
        "scope",
        "Harness scope",
        ["Harness Scope"],
        [
            ("Harness ID", 0),
            ("Rev", 1),
            ("Harness ID 2", 2),
            ("Rev 2", 3),
            ("Attribute", 4),
            ("Value", 5),
            ("Value 2", 6),
        ],
        group="Variance",
    ),
    View(
        "identity",
        "Harness identity",
        ["Harness name information"],
        [
            ("Harness ID", 0),
            ("Rev", 1),
            ("Date", 2),
            ("Field 4", 3),
            ("Source", 4),
            ("Harness ID 2", 5),
            ("Rev 2", 6),
            ("Date 2", 7),
            ("Site", 11),
            ("Drawing", 25),
        ],
        group="Variance",
    ),
]

VIEWS_BY_KEY = {v.key: v for v in VIEWS}
GROUP_ORDER = ["Components", "Wiring", "Structure", "Variance", "Other"]


def option_matrix(dsi: Dsi, option_start: int = 36) -> Tuple[List[str], List[List[str]]]:
    """Circuit x option-code presence grid, ``X`` where the code applies.

    Read from the raw trailing fields so only genuine codes are picked up,
    skipping the ``true``/``false`` flags that sit before them.
    """
    section = dsi.section("Harness circuit information")
    if section is None or not section.rows:
        return [], []
    known = {c for c in dsi.section("Composite Option Codes").column(0)} if "Composite Option Codes" in dsi else set()

    circuits = []
    ordered: List[str] = []
    seen = set()
    for row in section.rows:
        ref = section.field(row, 0)
        desc = section.field(row, 26)
        codes = set()
        for value in row[option_start:]:
            value = value.strip()
            if not value or value.lower() in ("true", "false"):
                continue
            codes.add(value)
            if value not in seen:
                seen.add(value)
                ordered.append(value)
        circuits.append((ref, desc, codes))

    # Prefer the declared order from Composite Option Codes, then any extras.
    if known:
        declared = [c for c in dsi.section("Composite Option Codes").column(0) if c in seen]
        extras = sorted(c for c in seen if c not in known)
        all_codes = declared + extras
    else:
        all_codes = sorted(ordered)

    headers = ["Circuit ID", "Description"] + all_codes
    rows = [
        [ref, desc] + ["X" if code in codes else "" for code in all_codes]
        for ref, desc, codes in circuits
    ]
    return headers, rows



def field_labels(section_name: str) -> Dict[int, str]:
    """Field index -> heading for a section, gathered from the views above.

    Single source of truth: the comparison report reuses the same names as the
    single-file views, so a column means the same thing in both outputs.
    """
    labels: Dict[int, str] = {}
    for view in VIEWS:
        if section_name not in view.section_names:
            continue
        for title, spec in view.columns:
            if isinstance(spec, int) and spec not in labels:
                labels[spec] = title
    return labels



# --------------------------------------------------------------------------- #
# Per-connector cavity charts
#
# One block per connector rather than one flat table: a title line, a row per
# wire landing in each cavity, and the connector part name underneath. This is
# the shape an engineer checks a connector against.
#
# Field mapping was confirmed against a known-good chart for connector 13D8A:
# every cell -- cavity, wire name, wire family, plating, far-end connector and
# cavity, option expression, and the cavity-5 plug -- reproduced exactly.
# --------------------------------------------------------------------------- #

CHART_COLUMNS = [
    "CAV",
    "NB FIL",
    "PART NUMBER WIRE",
    "MATERIAL",
    "FIN",
    "CAV FIN",
    "OPTION",
    "PLUG",
]

# Which component types get a chart. Splices carry cavities too, so add
# "SPLICE" here if you want charts for them as well.
CHART_TYPES = ("CONNECTOR", "IDC")

_OPERATORS = ((" + ", " && "), ("+", " && "), (" / ", " || "), ("/", " || "))


def format_option(expr: str) -> str:
    """Rewrite a DSI option expression in the notation used on the charts.

    ``+`` is AND and ``/`` is OR in the source; the charts show ``&&`` and
    ``||``. ``!`` (NOT) and parentheses carry through unchanged. Confirmed
    across 872 expressions whose tokens all resolve to real option codes.
    """
    if not expr:
        return ""
    out = expr.replace("+", " && ").replace("/", " || ")
    return " ".join(out.split())


def _cavity_sort_key(cav: str):
    """Natural order, so cavity 10 follows 9 rather than 1."""
    if not cav:
        return (2, (), "")
    chunks = re.findall(r"\d+|\D+", cav)
    key = tuple((0, int(c)) if c.isdigit() else (1, 0) for c in chunks)
    text = tuple(c.lower() for c in chunks if not c.isdigit())
    if cav.isdigit():
        return (0, (int(cav),), "")
    return (1, key + text, cav.lower())


class ConnectorChart:
    """One connector's block: heading, cavity rows, part name."""

    def __init__(self, ref: str, ctype: str, description: str, part_name: str,
                 part_number: str, cavities: str) -> None:
        self.ref = ref
        self.type = ctype
        self.description = description
        self.part_name = part_name
        self.part_number = part_number
        self.cavities = cavities
        self.rows: List[List[str]] = []


def connector_charts(dsi: Dsi, types: Sequence[str] = CHART_TYPES) -> List[ConnectorChart]:
    """Build a cavity chart for every component of the given types."""
    nodes = dsi.section("Harness main node components")
    if nodes is None:
        return []
    wires = dsi.find("Harness wire specification")
    plugs = dsi.find("plug", "Harness cavity plugs")
    terminals = dsi.find("terminals", "Harness terminals")
    wanted = {t.upper() for t in types}

    # Wire ends grouped by the connector they land on. One pass, in file order,
    # so rows inside a cavity keep the order the exporter wrote them.
    ends: Dict[str, List[Dict[str, str]]] = {}
    if wires is not None:
        for row in wires.rows:
            for near, far in ((8, 12), (12, 8)):
                ref = wires.field(row, near)
                if not ref:
                    continue
                near_cav, near_plate = (10, 11) if near == 8 else (14, 15)
                far_cav = 14 if near == 8 else 10
                ends.setdefault(ref, []).append(
                    {
                        "cav": wires.field(row, near_cav),
                        "wire": wires.field(row, 0),
                        "spec": wires.field(row, 28),
                        "material": wires.field(row, near_plate),
                        "fin": wires.field(row, far),
                        "cavfin": wires.field(row, far_cav),
                        "option": format_option(wires.field(row, 1)),
                    }
                )

    plug_by: Dict[str, Dict[str, str]] = {}
    if plugs is not None:
        for row in plugs.rows:
            plug_by.setdefault(plugs.field(row, 0), {})[plugs.field(row, 1)] = plugs.field(row, 8)

    term_cavs: Dict[str, set] = {}
    if terminals is not None:
        for row in terminals.rows:
            term_cavs.setdefault(terminals.field(row, 0), set()).add(terminals.field(row, 1))

    charts: List[ConnectorChart] = []
    for row in nodes.rows:
        ref = nodes.field(row, 0)
        ctype = nodes.field(row, 4).upper()
        if ctype not in wanted:
            continue
        chart = ConnectorChart(
            ref,
            nodes.field(row, 4),
            nodes.field(row, 6),
            nodes.field(row, 8),
            nodes.field(row, 12),
            nodes.field(row, 28),
        )

        my_ends = ends.get(ref, [])
        my_plugs = plug_by.get(ref, {})
        # Every cavity that has a wire, a plug, or a terminal. A plugged but
        # unwired cavity still belongs on the chart.
        universe = set(e["cav"] for e in my_ends) | set(my_plugs) | term_cavs.get(ref, set())
        universe.discard("")

        for cav in sorted(universe, key=_cavity_sort_key):
            at_cav = [e for e in my_ends if e["cav"] == cav]
            plug = my_plugs.get(cav, "")
            if not at_cav:
                chart.rows.append([cav, "", "", "", "", "", "", plug])
                continue
            for index, end in enumerate(at_cav):
                chart.rows.append(
                    [
                        cav,
                        end["wire"],
                        end["spec"],
                        end["material"],
                        end["fin"],
                        end["cavfin"],
                        end["option"],
                        plug if index == 0 else "",
                    ]
                )
        charts.append(chart)

    return charts
