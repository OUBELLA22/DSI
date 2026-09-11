#!/usr/bin/env python3
"""DSI Toolkit -- one entry point for both jobs.

Run it with no arguments and it asks what you want:

    1 file   -> pick which elements (connectors, splices, wires, ...) and get a
                clean Excel workbook, one sheet per element
    2 files  -> compare them and get a difference report

Or drive it from the command line::

    python3 dsi_tool.py view    HARNESS.dsi
    python3 dsi_tool.py view    HARNESS.dsi --elements connectors,splices,wires
    python3 dsi_tool.py compare OLD.dsi NEW.dsi
    python3 dsi_tool.py compare OLD.dsi NEW.dsi --all-rows -o REPORT.xlsx

Standard library only, so it runs on a stock Python with nothing installed.
"""

from __future__ import annotations

import glob
import os
import re
import sys
from typing import Dict, List, Optional, Sequence, Tuple

import compare as cmp_mod
import views as views_mod
import xlsx
from dsi import Dsi, DsiError

BANNER = "DSI Toolkit  --  Capital HarnessXC / ModularXC"
LINE = "=" * 62

# Verdict -> fill used for that row's value cells.
VERDICT_FILL = {
    cmp_mod.MODIFY: xlsx.YELLOW,
    cmp_mod.ADD: xlsx.GREEN,
    cmp_mod.REMOVE: xlsx.RED,
}


# --------------------------------------------------------------------------- #
# Metadata
# --------------------------------------------------------------------------- #
def file_metadata(dsi: Dsi) -> List[Tuple[str, str]]:
    """Pull ``Key: value`` pairs out of the export banner."""
    out = []
    for line in dsi.preamble:
        text = line.lstrip(dsi.comment).strip()
        if text.startswith("*"):
            text = text.lstrip("*").strip()
        if not text or set(text) == {"*"}:
            continue
        match = re.match(r"([^:=]+?)\s*[:=]\s+(.+)$", text)
        if match:
            key = match.group(1).strip()
            value = match.group(2).strip()
            if key and value:
                out.append((key, value))
    return out


# --------------------------------------------------------------------------- #
# Single file: element views
# --------------------------------------------------------------------------- #
MATRIX_KEY = "matrix"
CHART_KEY = "charts"


def available_views(dsi: Dsi) -> List[Tuple[str, str, int]]:
    """``(key, label, row count)`` for every view with data in this file."""
    out = []
    for view in views_mod.VIEWS:
        count = view.count(dsi)
        if count:
            out.append((view.key, view.label, count))
        if view.key == "connectors" and count:
            charts = sum(
                1
                for r in dsi["Harness main node components"].rows
                if dsi["Harness main node components"].field(r, 4).upper()
                in views_mod.CHART_TYPES
            )
            out.append((CHART_KEY, "Connector cavity charts", charts))
    if "Harness circuit information" in dsi:
        circuits = dsi["Harness circuit information"]
        if circuits.rows:
            out.append((MATRIX_KEY, "Option matrix (circuit x code)", len(circuits.rows)))
    return out


def build_chart_sheet(dsi: Dsi) -> Tuple[Optional[xlsx.Sheet], int]:
    """Stacked per-connector cavity charts, one block each.

    Row 1 is a block title rather than a column header, so the sheet is written
    with header_style=None and every row driven through add().
    """
    charts = views_mod.connector_charts(dsi)
    if not charts:
        return None, 0

    cols = views_mod.CHART_COLUMNS
    width = len(cols)
    first = charts[0]

    def title_row(chart: views_mod.ConnectorChart) -> Tuple[List[str], Dict[int, str]]:
        row = [""] * width
        row[0] = chart.ref
        row[6] = chart.description
        return row, {0: xlsx.BOLD, 6: xlsx.BOLD}

    head, head_fills = title_row(first)
    sheet = xlsx.Sheet(
        "Connector Charts",
        head,
        table=False,
        header_style=None,
        header_fills=head_fills,
        center_cols={0, 3, 5},
        widths={0: 6, 1: 14, 2: 20, 3: 10, 4: 10, 5: 9, 6: 42, 7: 14},
    )

    def emit_block(chart: views_mod.ConnectorChart) -> None:
        sheet.add(list(cols), fills={i: xlsx.HEADER for i in range(width)})
        for row in chart.rows:
            sheet.add(row)
        footer = [""] * width
        footer[0] = chart.part_name
        footer[1] = chart.part_number
        footer[6] = "cavities: {}".format(chart.cavities) if chart.cavities else ""
        sheet.add(footer, fills={0: xlsx.BOLD})
        sheet.add([""] * width)

    emit_block(first)
    for chart in charts[1:]:
        row, fills = title_row(chart)
        sheet.add(row, fills=fills)
        emit_block(chart)

    return sheet, len(charts)


def build_view_workbook(dsi: Dsi, keys: Sequence[str], path: str) -> List[Tuple[str, int, int]]:
    """Write the element workbook. Returns ``(sheet, cols, rows)`` per sheet."""
    ctx = views_mod.Ctx(dsi)
    sheets: List[xlsx.Sheet] = []
    summary: List[Tuple[str, int, int]] = []

    overview = xlsx.Sheet("Overview", ["Property", "Value"], table=False, widths={0: 30, 1: 80})
    for key, value in file_metadata(dsi):
        overview.add([key, value])
    overview.add(["", ""])
    overview.add(["Sections in file", str(len(dsi.sections))], fills={0: xlsx.BOLD})
    overview.add(["Data rows in file", str(dsi.row_count)], fills={0: xlsx.BOLD})
    overview.add(["", ""])
    overview.add(["Element", "Rows"], fills={0: xlsx.BOLD, 1: xlsx.BOLD})
    sheets.append(overview)

    for key in keys:
        if key == CHART_KEY:
            sheet, n = build_chart_sheet(dsi)
            if sheet is None:
                continue
            sheets.append(sheet)
            overview.add(["Connector cavity charts", "{} connectors".format(n)])
            summary.append(("Connector Charts", len(views_mod.CHART_COLUMNS), len(sheet)))
            continue

        if key == MATRIX_KEY:
            headers, rows = views_mod.option_matrix(dsi)
            if not headers:
                continue
            centre = set(range(2, len(headers)))
            sheet = xlsx.Sheet(
                "Option Matrix",
                headers,
                text_cols=set(range(len(headers))),
                center_cols=centre,
                widths={i: 7.0 for i in centre},
            )
            for row in rows:
                sheet.add(row)
            sheets.append(sheet)
            overview.add(["Option matrix (circuit x code)", str(len(rows))])
            summary.append(("Option Matrix", len(headers), len(rows)))
            continue

        view = views_mod.VIEWS_BY_KEY.get(key)
        if view is None:
            continue
        headers, rows = view.build(dsi, ctx)
        if not rows:
            continue
        sheet = xlsx.Sheet(view.sheet, headers, text_cols=view.text_columns())
        for row in rows:
            sheet.add(row)
        sheets.append(sheet)
        overview.add([view.label, str(len(rows))])
        summary.append((view.sheet, len(headers), len(rows)))

    if len(sheets) == 1:
        raise DsiError("nothing to write: the selected elements have no rows in this file")

    xlsx.write_workbook(path, sheets)
    return summary


# --------------------------------------------------------------------------- #
# Two files: comparison report
# --------------------------------------------------------------------------- #
def build_compare_workbook(
    dsi_a: Dsi,
    dsi_b: Dsi,
    results: Sequence[cmp_mod.SectionResult],
    path: str,
    include_unchanged: bool = False,
) -> None:
    sheets: List[xlsx.Sheet] = []

    # -- SYNTHESIS ----------------------------------------------------------
    synthesis = xlsx.Sheet(
        "SYNTHESIS",
        [
            "Section",
            "Rows file 1",
            "Rows file 2",
            "OK",
            "MODIFY",
            "ADD",
            "REMOVE",
            "REVERSED",
            "Match key",
            "Ignored fields",
            "Why this key",
        ],
        center_cols={1, 2, 3, 4, 5, 6, 7},
        widths={0: 34, 8: 34, 9: 16, 10: 60},
    )
    for result in results:
        modify = result.count(cmp_mod.MODIFY)
        added = result.count(cmp_mod.ADD)
        removed = result.count(cmp_mod.REMOVE)
        fills = {}
        if modify:
            fills[4] = xlsx.YELLOW
        if added:
            fills[5] = xlsx.GREEN
        if removed:
            fills[6] = xlsx.RED
        ignored = ", ".join("f{}".format(i) for i in sorted(result.rule.ignore)) or "-"
        note = result.rule.note or ("" if result.specific else "no rule defined; fell back to field 0")
        synthesis.add(
            [
                result.name,
                result.rows_a,
                result.rows_b,
                result.count(cmp_mod.OK),
                modify,
                added,
                removed,
                result.reversed_count or "",
                result.rule.describe_key(),
                ignored,
                note,
            ],
            fills=fills,
        )
    sheets.append(synthesis)

    # -- DIFFERENCES (long format, only the deltas) -------------------------
    diffs = xlsx.Sheet(
        "DIFFERENCES",
        ["Section", "Check", "Match", "Key", "Field", "Value in file 1", "Value in file 2"],
        text_cols={3, 5, 6},
        center_cols={1, 2},
        widths={0: 30, 1: 10, 2: 10, 3: 28, 4: 30, 5: 36, 6: 36},
    )
    for result in results:
        for pair in result.pairs:
            if pair.verdict == cmp_mod.OK and not pair.reversed_match:
                continue
            key_text = " | ".join(pair.key)
            if pair.verdict == cmp_mod.OK:
                # Reversed but otherwise identical: no field differs, yet the
                # file did change direction, so it still belongs in the report.
                diffs.add(
                    [result.name, cmp_mod.OK, "reversed", key_text,
                     "(direction reversed, no other change)", "", ""],
                    fills={2: xlsx.YELLOW},
                )
                continue
            match_text = "reversed" if pair.reversed_match else "direct"
            if pair.verdict == cmp_mod.MODIFY:
                for index_a, index_b, value_a, value_b in pair.diffs:
                    field = cmp_mod.field_header(result.name, index_a)
                    if index_a != index_b:
                        field += " vs " + cmp_mod.field_header(result.name, index_b)
                    diffs.add(
                        [result.name, cmp_mod.MODIFY, match_text, key_text, field, value_a, value_b],
                        fills={1: xlsx.YELLOW, 5: xlsx.YELLOW, 6: xlsx.YELLOW},
                    )
            elif pair.verdict == cmp_mod.ADD:
                diffs.add(
                    [result.name, cmp_mod.ADD, "", key_text, "(whole row)", "", ":".join(pair.row_b or [])],
                    fills={1: xlsx.GREEN, 6: xlsx.GREEN},
                )
            else:
                diffs.add(
                    [result.name, cmp_mod.REMOVE, "", key_text, "(whole row)", ":".join(pair.row_a or []), ""],
                    fills={1: xlsx.RED, 5: xlsx.RED},
                )
    sheets.append(diffs)

    # -- one sheet per section, side by side --------------------------------
    for result in results:
        rows_to_show = [
            p
            for p in result.pairs
            if include_unchanged or p.verdict != cmp_mod.OK or p.reversed_match
        ]
        if not rows_to_show:
            continue
        width = result.width
        headers = ["CHECK", "MATCH", "KEY"]
        for suffix in ("(1)", "(2)"):
            for u in range(width):
                headers.append("{} {}".format(cmp_mod.field_header(result.name, u), suffix))
        sheet = xlsx.Sheet(
            result.name[:31],
            headers,
            text_cols=set(range(2, len(headers))),
            center_cols={0, 1},
            widths={0: 11, 1: 10, 2: 26},
        )
        offset_a = 3
        offset_b = 3 + width
        for pair in rows_to_show:
            row: List[str] = [
                pair.verdict,
                "reversed" if pair.reversed_match else "",
                " | ".join(pair.key),
            ]
            row += [(pair.row_a[u] if pair.row_a and u < len(pair.row_a) else "") for u in range(width)]
            row += [(pair.row_b[u] if pair.row_b and u < len(pair.row_b) else "") for u in range(width)]
            fills: Dict[int, str] = {}
            verdict_fill = VERDICT_FILL.get(pair.verdict)
            if verdict_fill:
                fills[0] = verdict_fill
            for u in sorted(result.rule.ignore):
                if u < width:
                    fills[offset_a + u] = xlsx.GREY
                    fills[offset_b + u] = xlsx.GREY
            if pair.verdict == cmp_mod.MODIFY:
                for index_a, index_b, _, _ in pair.diffs:
                    fills[offset_a + index_a] = xlsx.YELLOW
                    fills[offset_b + index_b] = xlsx.YELLOW
            elif pair.verdict == cmp_mod.ADD:
                for u in range(width):
                    if (pair.row_b[u] if pair.row_b and u < len(pair.row_b) else ""):
                        fills[offset_b + u] = xlsx.GREEN
            elif pair.verdict == cmp_mod.REMOVE:
                for u in range(width):
                    if (pair.row_a[u] if pair.row_a and u < len(pair.row_a) else ""):
                        fills[offset_a + u] = xlsx.RED
            sheet.add(row, fills=fills)
        sheets.append(sheet)

    # -- FILES --------------------------------------------------------------
    files = xlsx.Sheet("FILES", ["", "File 1", "File 2"], table=False, widths={0: 28, 1: 60, 2: 60})
    files.add(["Path", dsi_a.path or "", dsi_b.path or ""], fills={0: xlsx.BOLD})
    meta_a = dict(file_metadata(dsi_a))
    meta_b = dict(file_metadata(dsi_b))
    for key in sorted(set(meta_a) | set(meta_b)):
        files.add([key, meta_a.get(key, ""), meta_b.get(key, "")], fills={0: xlsx.BOLD})
    files.add(["Sections", str(len(dsi_a.sections)), str(len(dsi_b.sections))], fills={0: xlsx.BOLD})
    files.add(["Data rows", str(dsi_a.row_count), str(dsi_b.row_count)], fills={0: xlsx.BOLD})
    sheets.append(files)

    xlsx.write_workbook(path, sheets)


def print_synthesis(results: Sequence[cmp_mod.SectionResult]) -> None:
    """Console version of the synthesis, since the report is a file."""
    header = ("SECTION", "F1", "F2", "OK", "MOD", "ADD", "REM", "REV")
    rows = [
        (
            r.name[:38],
            str(r.rows_a),
            str(r.rows_b),
            str(r.count(cmp_mod.OK)),
            str(r.count(cmp_mod.MODIFY)),
            str(r.count(cmp_mod.ADD)),
            str(r.count(cmp_mod.REMOVE)),
            str(r.reversed_count or ""),
        )
        for r in results
    ]
    widths = [max(len(h), *(len(r[i]) for r in rows)) if rows else len(h) for i, h in enumerate(header)]
    fmt = "  ".join("{{:<{}}}".format(w) if i == 0 else "{{:>{}}}".format(w) for i, w in enumerate(widths))
    print(fmt.format(*header))
    print("  ".join("-" * w for w in widths))
    for row in rows:
        print(fmt.format(*row))
    total_mod = sum(r.count(cmp_mod.MODIFY) for r in results)
    total_add = sum(r.count(cmp_mod.ADD) for r in results)
    total_rem = sum(r.count(cmp_mod.REMOVE) for r in results)
    print("  ".join("-" * w for w in widths))
    print("{} section(s):  {} modified, {} added, {} removed".format(
        len(results), total_mod, total_add, total_rem))
    unruled = [r.name for r in results if not r.specific]
    if unruled:
        print()
        print("No match rule for these, fell back to field 0 -- check the key on SYNTHESIS:")
        for name in unruled:
            print("   -", name)


# --------------------------------------------------------------------------- #
# Prompting
# --------------------------------------------------------------------------- #
def ask(prompt: str, default: str = "") -> str:
    try:
        answer = input(prompt).strip()
    except EOFError:
        return default
    return answer or default


def find_dsi_files() -> List[str]:
    found = []
    for pattern in ("*.dsi", "*.DSI", "*.Dsi"):
        found.extend(glob.glob(pattern))
    unique = sorted({os.path.normpath(p) for p in found})
    return unique


def choose_file(role: str, already: Optional[str] = None) -> Optional[str]:
    """Prompt for a .dsi path.

    ``already`` is only annotated, never removed from the list: dropping it
    would renumber everything between the two prompts, so the number the user
    just read would mean something else. Picking the same file twice is allowed
    -- comparing a file against itself is a useful way to check the tool.
    """
    candidates = find_dsi_files()
    print()
    if candidates:
        print("DSI files in this folder:")
        for i, path in enumerate(candidates, 1):
            size = os.path.getsize(path) / 1024.0
            mark = "   <- chosen as file 1" if already and path == already else ""
            print("  {:>2}  {}  ({:.0f} KB){}".format(i, path, size, mark))
        print("   0  type a different path")
        while True:
            answer = ask("\n{} -- number or path: ".format(role))
            if not answer:
                return None
            if answer.isdigit():
                index = int(answer)
                if index == 0:
                    break
                if 1 <= index <= len(candidates):
                    return candidates[index - 1]
                print("  Pick 0-{}.".format(len(candidates)))
                continue
            if os.path.isfile(answer):
                return answer
            print("  No such file: {}".format(answer))
    while True:
        answer = ask("\n{} -- path to .dsi file: ".format(role))
        if not answer:
            return None
        if os.path.isfile(answer):
            return answer
        print("  No such file: {}".format(answer))


def parse_selection(text: str, count: int) -> Optional[List[int]]:
    """Parse ``1,3,5``, ``1-4``, ``a`` into 1-based indices. None if invalid."""
    text = text.strip().lower()
    if text in ("a", "all", "*"):
        return list(range(1, count + 1))
    if not text:
        return None
    picked: List[int] = []
    for part in text.replace(" ", "").split(","):
        if not part:
            continue
        try:
            if "-" in part:
                low, high = part.split("-", 1)
                picked.extend(range(int(low), int(high) + 1))
            else:
                picked.append(int(part))
        except ValueError:
            return None
    seen = set()
    ordered = []
    for index in picked:
        if 1 <= index <= count and index not in seen:
            seen.add(index)
            ordered.append(index)
    return ordered or None


def choose_elements(dsi: Dsi) -> Optional[List[str]]:
    entries = available_views(dsi)
    if not entries:
        print("\nThis file has no rows in any known section.")
        return None

    by_group: Dict[str, List[Tuple[int, str, int]]] = {}
    numbering: List[str] = []
    for key, label, count in entries:
        view = views_mod.VIEWS_BY_KEY.get(key)
        if key == CHART_KEY:
            group = "Components"
        else:
            group = view.group if view else "Variance"
        numbering.append(key)
        by_group.setdefault(group, []).append((len(numbering), label, count))

    print()
    print("What do you want to see?")
    for group in views_mod.GROUP_ORDER:
        if group not in by_group:
            continue
        print("\n  {}".format(group.upper()))
        for index, label, count in by_group[group]:
            print("    {:>2}  {:<34} {:>6} rows".format(index, label, count))
    print("\n     a  everything listed above")
    print("     q  cancel")

    while True:
        answer = ask("\nSelect (e.g. 1,3,8  or  1-5  or  a): ")
        if answer.lower() in ("q", "quit"):
            return None
        picked = parse_selection(answer, len(numbering))
        if picked:
            return [numbering[i - 1] for i in picked]
        print("  Didn't understand that. Use numbers, ranges, or 'a'.")


def choose_sections(dsi_a: Dsi, dsi_b: Dsi) -> Optional[List[str]]:
    names: List[str] = []
    for source in (dsi_a, dsi_b):
        for section in source.sections:
            if section.name in cmp_mod.SKIP_SECTIONS or section.name in names:
                continue
            if section.rows:
                names.append(section.name)
    print()
    print("Sections with data:")
    for i, name in enumerate(names, 1):
        a = dsi_a.section(name)
        b = dsi_b.section(name)
        print("  {:>2}  {:<44} {:>5} / {:>5}".format(
            i, name[:44], len(a.rows) if a else 0, len(b.rows) if b else 0))
    print("\n     a  compare all of them")
    while True:
        answer = ask("\nSelect (e.g. 1,3,8  or  a): ", "a")
        picked = parse_selection(answer, len(names))
        if picked:
            return [names[i - 1] for i in picked]
        print("  Didn't understand that.")


def load(path: str) -> Optional[Dsi]:
    try:
        dsi = Dsi.read(path)
    except (OSError, DsiError, UnicodeDecodeError) as exc:
        print("\nCould not read {}:\n  {}".format(path, exc))
        return None
    print("  loaded {}  ->  {} sections, {} rows".format(
        os.path.basename(path), len(dsi.sections), dsi.row_count))
    if dsi.mixed_newlines:
        print("  note: this file has mixed line endings")
    return dsi


def default_output(*paths: str) -> str:
    stems = [os.path.splitext(os.path.basename(p))[0] for p in paths]
    if len(stems) == 1:
        return "{}_VIEW.xlsx".format(stems[0])
    return "COMPARE_{}_vs_{}.xlsx".format(stems[0][:24], stems[1][:24])


def ask_output(default: str) -> str:
    answer = ask("\nOutput .xlsx [{}]: ".format(default), default)
    if not answer.lower().endswith(".xlsx"):
        answer += ".xlsx"
    return answer


# --------------------------------------------------------------------------- #
# Flows
# --------------------------------------------------------------------------- #
def flow_view(path: Optional[str] = None, keys: Optional[Sequence[str]] = None,
              out: Optional[str] = None) -> int:
    if path is None:
        path = choose_file("DSI file")
        if path is None:
            return 1
    print()
    dsi = load(path)
    if dsi is None:
        return 1

    if keys is None:
        keys = choose_elements(dsi)
        if not keys:
            print("\nCancelled.")
            return 1
    else:
        known = set(views_mod.VIEWS_BY_KEY) | {MATRIX_KEY, CHART_KEY}
        bad = [k for k in keys if k not in known]
        if bad:
            print("Unknown element(s): {}".format(", ".join(bad)))
            print("Known: {}".format(", ".join(sorted(known))))
            return 2

    target = out or ask_output(default_output(path))
    try:
        summary = build_view_workbook(dsi, list(keys), target)
    except DsiError as exc:
        print("\n{}".format(exc))
        return 1

    print("\nWrote {}".format(target))
    for name, cols, rows in summary:
        print("  {:<32} {:>4} cols  {:>6} rows".format(name, cols, rows))
    return 0


def flow_compare(path_a: Optional[str] = None, path_b: Optional[str] = None,
                 out: Optional[str] = None, all_rows: bool = False,
                 interactive: bool = True) -> int:
    if path_a is None:
        path_a = choose_file("File 1 (the reference / older one)")
        if path_a is None:
            return 1
    if path_b is None:
        path_b = choose_file("File 2 (the one to check)", already=path_a)
        if path_b is None:
            return 1
    if os.path.abspath(path_a) == os.path.abspath(path_b):
        print("\nNote: both files are the same, so everything should come back OK.")
    print()
    dsi_a = load(path_a)
    dsi_b = load(path_b)
    if dsi_a is None or dsi_b is None:
        return 1

    only = None
    if interactive:
        answer = ask("\nCompare every section? [Y/n]: ", "y").lower()
        if answer.startswith("n"):
            only = choose_sections(dsi_a, dsi_b)
            if not only:
                return 1
        answer = ask("Include unchanged (OK) rows in the section sheets? [y/N]: ", "n").lower()
        all_rows = answer.startswith("y")

    results = cmp_mod.compare(dsi_a, dsi_b, only)
    if not results:
        print("\nNothing to compare: no section has rows in either file.")
        return 1

    target = out or (ask_output(default_output(path_a, path_b)) if interactive
                     else default_output(path_a, path_b))
    build_compare_workbook(dsi_a, dsi_b, results, target, include_unchanged=all_rows)

    print()
    print(LINE)
    print("COMPARISON SYNTHESIS")
    print(LINE)
    print_synthesis(results)
    print()
    print("Wrote {}".format(target))
    print("  SYNTHESIS    per-section counts, and the match key used for each")
    print("  DIFFERENCES  every delta as one row: section, key, field, both values")
    print("  <section>    side by side, file 1 block then file 2 block")
    if not all_rows:
        print("               (unchanged rows omitted; rerun with --all-rows to keep them)")
    return 0


def interactive() -> int:
    print()
    print(LINE)
    print("  " + BANNER)
    print(LINE)
    print()
    print("How many DSI files are you working with?")
    print()
    print("  1   One file    ->  pick the elements you want to see (connectors,")
    print("                      splices, wires, ...) and get a clean workbook")
    print("  2   Two files   ->  compare them and report every difference")
    print("  q   Quit")
    while True:
        answer = ask("\nChoice [1/2/q]: ").lower()
        if answer in ("q", "quit", "exit"):
            return 0
        if answer == "1":
            return flow_view()
        if answer == "2":
            return flow_compare()
        print("  Enter 1, 2 or q.")


# --------------------------------------------------------------------------- #
# Command line
# --------------------------------------------------------------------------- #
USAGE = """DSI Toolkit

  python3 dsi_tool.py                          interactive: asks 1 file or 2
  python3 dsi_tool.py view FILE.dsi [options]  one file  -> element workbook
  python3 dsi_tool.py compare A.dsi B.dsi      two files -> difference report

Options
  -o, --out PATH        output .xlsx (default is derived from the input names)
  --elements a,b,c      view mode: element keys, skips the menu
  --all-rows            compare mode: keep unchanged rows in section sheets
  --list-elements       print the element keys and exit
"""


def main(argv: Sequence[str]) -> int:
    args = list(argv[1:])
    if not args:
        try:
            return interactive()
        except KeyboardInterrupt:
            print("\nCancelled.")
            return 1

    if args[0] in ("-h", "--help", "help"):
        print(USAGE)
        return 0

    if args[0] == "--list-elements":
        for view in views_mod.VIEWS:
            print("  {:<12} {:<34} {}".format(view.key, view.label, view.section_names[0]))
        print("  {:<12} {:<34} {}".format(CHART_KEY, "Connector cavity charts",
                                          "Harness main node components"))
        print("  {:<12} {}".format(MATRIX_KEY, "Option matrix (circuit x code)"))
        return 0

    mode = args.pop(0)
    out: Optional[str] = None
    elements: Optional[List[str]] = None
    all_rows = False
    positional: List[str] = []
    while args:
        arg = args.pop(0)
        if arg in ("-o", "--out"):
            if not args:
                print("--out needs a path")
                return 2
            out = args.pop(0)
        elif arg == "--elements":
            if not args:
                print("--elements needs a comma-separated list")
                return 2
            elements = [e.strip() for e in args.pop(0).split(",") if e.strip()]
        elif arg == "--all-rows":
            all_rows = True
        elif arg.startswith("-"):
            print("Unknown option {}".format(arg))
            return 2
        else:
            positional.append(arg)

    if mode == "view":
        if len(positional) != 1:
            print(USAGE)
            return 2
        return flow_view(positional[0], elements, out)
    if mode == "compare":
        if len(positional) != 2:
            print(USAGE)
            return 2
        return flow_compare(positional[0], positional[1], out, all_rows, interactive=False)

    print(USAGE)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
