#!/usr/bin/env python3
"""Compare two DSI exports section by section.

The whole of the VBA comparison toolkit is roughly sixty near-identical
subroutines. They differ only in three things: which fields identify a row,
which fields to skip, and whether the two ends of a segment may be swapped.
Everything else -- the nested loop, the claimed-row guard, the OK/MODIFY/ADD/
REMOVE verdict, the colouring -- is copied text. So the knowledge lives in the
:data:`RULES` table below and there is one comparison function.

Matching mirrors the VBA's semantics deliberately:

* a row in file 2 can be claimed by only one row in file 1;
* the first unclaimed row with a matching key wins, so duplicate keys pair up
  in file order rather than erroring;

but it runs in one pass over a dict instead of a nested scan, so it is O(n)
rather than O(n squared).

Where the VBA's exact field indices could not be recovered with confidence, the
rule carries a ``note`` explaining the choice, and every rule actually used is
printed on the SYNTHESIS sheet. Compare that against the VBA's own output and
edit :data:`RULES` if a key needs adjusting -- that is one line here instead of
a hunt through sixty subroutines.
"""

from __future__ import annotations

from collections import deque
from typing import Dict, FrozenSet, Iterable, List, Optional, Sequence, Tuple

from dsi import Dsi, Section
import views

OK = "OK"
MODIFY = "MODIFY"
ADD = "ADD"
REMOVE = "REMOVE"


class Rule:
    """How to match rows of one section between two files.

    Args:
        sections: candidate section names. The two engine families spell the
            same section differently (``terminals`` vs ``Harness terminals``),
            so both are listed and whichever is present is used.
        keys: field-index tuples tried in order. A later tuple is a fallback
            for rows the earlier one could not pair.
        swap: field pairs that exchange places when a row is recorded in the
            opposite direction, for sections whose rows describe a segment
            between two nodes. A branch stored A->B in one file and B->A in the
            other is the same branch, so it must not surface as one ADD plus
            one REMOVE.

            The swap applies to the key *and* to the field comparison. Giving
            branch configuration ``((0, 3), (2, 5))`` says: end nodes 0 and 3
            trade places, and so do their coordinates 2 and 5. Without the
            second pair a reversed branch would match and then report its own
            node fields as differences, which is noise.
        ignore: fields excluded from the cell-by-cell comparison. Used for
            values that change without meaning a design change.
        note: why this key was chosen. Printed on the SYNTHESIS sheet.
    """

    def __init__(
        self,
        sections: Sequence[str],
        keys: Sequence[Sequence[int]] = ((0,),),
        swap: Sequence[Sequence[int]] = (),
        ignore: Iterable[int] = (),
        note: str = "",
    ) -> None:
        self.sections = tuple(sections)
        self.keys: Tuple[Tuple[int, ...], ...] = tuple(tuple(k) for k in keys)
        self.swap: Tuple[Tuple[int, int], ...] = tuple((a, b) for a, b in swap)
        self.swap_map: Dict[int, int] = {}
        for first, second in self.swap:
            self.swap_map[first] = second
            self.swap_map[second] = first
        self.ignore: FrozenSet[int] = frozenset(ignore)
        self.note = note

    def describe_key(self) -> str:
        parts = [" + ".join("f{}".format(f) for f in k) for k in self.keys]
        text = "  then  ".join(parts)
        if self.swap:
            pairs = ", ".join("f{}<->f{}".format(a, b) for a, b in self.swap)
            text += "   (may reverse: {})".format(pairs)
        return text


# --------------------------------------------------------------------------- #
# The match-key table: the domain knowledge, in one place.
# --------------------------------------------------------------------------- #
RULES: List[Rule] = [
    Rule(["Harness name information"], [(0,)], note="one row; matched on harness id"),
    Rule(["Harness circuit information"], [(0,)], note="circuit id is unique per row"),
    Rule(
        ["Harness branch configuration"],
        [(0, 3)],
        swap=[(0, 3), (2, 5)],
        ignore={2, 5, 19, 20, 21, 22, 23},
        note="segment end nodes. Coordinate fields ignored: geometry shifts "
        "without a design change and would swamp the report. Remove them from "
        "ignore to diff geometry.",
    ),
    Rule(["Harness wire specification"], [(0,)], note="wire name is unique"),
    Rule(["Harness main node components"], [(0,)], note="component reference is unique"),
    Rule(["terminals", "Harness terminals"], [(0, 1)], note="connector reference + cavity"),
    Rule(["seals", "Harness cavity seals"], [(0, 1)], note="connector reference + cavity"),
    Rule(["plug", "Harness cavity plugs"], [(0, 1)], note="connector reference + cavity"),
    Rule(["clips"], [(0,)], note="clip reference is unique"),
    Rule(["grommets"], [(0,)], note="grommet reference is unique"),
    Rule(["other components"], [(0,)], note="reference is unique"),
    Rule(
        ["Harness extra node components", "Extra_Node_Component"],
        [(0, 8), (0,)],
        note="reference + part name. Reference alone repeats (28 distinct "
        "across 49 rows in the sample), so the part name disambiguates.",
    ),
    Rule(
        ["Harness branch insulations", "Branch_insulation"],
        [(0, 2)],
        swap=[(0, 2)],
        ignore={4},
        note="the two end nodes. Field 4 is a per-export row sequence "
        "(1..n, renumbers on every export) so it is ignored.",
    ),
    Rule(["Harness multicores"], [(0,)], note="multicore id is unique"),
    Rule(["Harness center strips"], [(0, 1)]),
    Rule(["Module child details"], [(7,)], ignore={4, 9}),
    Rule(["Module compatibility details"], [(0, 1)], swap=[(0, 1)]),
    Rule(["Manual BOM quantities"], [(0,)]),
    Rule(["Harness wire through nodes"], [(0, 2)], note="wire name + sequence"),
    Rule(["Harness branch insulation through nodes"], [(0, 1)], note="insulation + sequence"),
    Rule(["Harness mid wire components"], [(0, 1)]),
    Rule(["Harness wire / multicore markers"], [(0, 1)]),
    Rule(["Harness pin mappings"], [(0, 1)]),
    Rule(
        ["Harness Scope"],
        [(0, 1, 2, 4), (0, 1, 2)],
        note="harness id + rev + attribute name. The attribute is added to the "
        "VBA key because each harness carries three attribute rows that would "
        "otherwise pair arbitrarily.",
    ),
    Rule(["Composite Option Codes"], [(0,)], note="option code is unique"),
    Rule(["Harness note information"], [(0,)]),
    # Sections the other engine family emits.
    Rule(["Assemblies", "Harness assembly items"], [(8, 26), (8,)]),
    Rule(["Multi Location Components", "Harness multiple location components"], [(0, 3, 8), (0, 3)]),
    Rule(["property"], [(0, 1, 2, 3), (0, 1, 2)]),
    Rule(["tape"], [(0,)]),
    Rule(["extracavity"], [(0, 1)]),
    Rule(["WIREend"], [(0, 1)]),
]

DEFAULT_RULE = Rule([], [(0,)], note="no rule defined; fell back to field 0")

# Not data, so never compared.
SKIP_SECTIONS = {"End of file marker."}


def rule_for(section_name: str) -> Tuple[Rule, bool]:
    """Return ``(rule, is_specific)`` for a section name."""
    for rule in RULES:
        if section_name in rule.sections:
            return rule, True
    return DEFAULT_RULE, False


# --------------------------------------------------------------------------- #
# Comparison
# --------------------------------------------------------------------------- #


class Pair:
    """One matched or unmatched row pair and its verdict.

    ``diffs`` holds ``(field_a, field_b, value_a, value_b)``. The two indices
    differ only for a reversed match, where file 1 field 0 is compared against
    file 2 field 3.
    """

    __slots__ = ("verdict", "key", "row_a", "row_b", "diffs", "reversed_match")

    def __init__(
        self,
        verdict: str,
        key: Tuple[str, ...],
        row_a: Optional[List[str]],
        row_b: Optional[List[str]],
        diffs: Optional[List[Tuple[int, int, str, str]]] = None,
        reversed_match: bool = False,
    ) -> None:
        self.verdict = verdict
        self.key = key
        self.row_a = row_a
        self.row_b = row_b
        self.diffs = diffs or []
        self.reversed_match = reversed_match


class SectionResult:
    """Outcome for one section."""

    def __init__(self, name: str, rule: Rule, specific: bool, key_used: Tuple[int, ...]) -> None:
        self.name = name
        self.rule = rule
        self.specific = specific
        self.key_used = key_used
        self.pairs: List[Pair] = []
        self.rows_a = 0
        self.rows_b = 0
        self.width = 0
        self.only_in: Optional[str] = None  # 'a' or 'b' when the section is one-sided

    def count(self, verdict: str) -> int:
        return sum(1 for p in self.pairs if p.verdict == verdict)

    @property
    def reversed_count(self) -> int:
        """Rows that only paired once the segment direction was flipped."""
        return sum(1 for p in self.pairs if p.reversed_match)

    @property
    def changed(self) -> int:
        return sum(1 for p in self.pairs if p.verdict != OK)


def _key_of(
    section: Section,
    row: Sequence[str],
    fields: Sequence[int],
    swap_map: Optional[Dict[int, int]] = None,
) -> Tuple[str, ...]:
    if swap_map:
        fields = [swap_map.get(f, f) for f in fields]
    return tuple(section.field(row, f) for f in fields)


class _Index:
    """key -> queue of row indices, handing out each row at most once."""

    def __init__(self, section: Section, fields: Sequence[int]) -> None:
        self.buckets: Dict[Tuple[str, ...], deque] = {}
        for index, row in enumerate(section.rows):
            key = _key_of(section, row, fields, None)
            self.buckets.setdefault(key, deque()).append(index)

    def take(self, key: Tuple[str, ...], claimed: set) -> Optional[int]:
        bucket = self.buckets.get(key)
        while bucket:
            index = bucket.popleft()
            if index not in claimed:
                return index
        return None


def compare_section(name: str, section_a: Optional[Section], section_b: Optional[Section]) -> SectionResult:
    """Diff one section. Either side may be ``None`` when only one file has it."""
    rule, specific = rule_for(name)
    result = SectionResult(name, rule, specific, rule.keys[0])
    result.rows_a = len(section_a.rows) if section_a else 0
    result.rows_b = len(section_b.rows) if section_b else 0
    result.width = max(
        section_a.width if section_a else 0,
        section_b.width if section_b else 0,
    )

    if section_a is None and section_b is None:
        return result
    if section_b is None:
        result.only_in = "a"
        for row in section_a.rows:
            result.pairs.append(Pair(REMOVE, _key_of(section_a, row, rule.keys[0]), list(row), None))
        return result
    if section_a is None:
        result.only_in = "b"
        for row in section_b.rows:
            result.pairs.append(Pair(ADD, _key_of(section_b, row, rule.keys[0]), None, list(row)))
        return result

    # Ordered passes: primary key direct, primary key reversed, then each
    # fallback the same way. Full passes rather than choosing a variant per row,
    # so a direct primary match always takes precedence over a reversed or
    # fallback one.
    passes: List[Tuple[Tuple[int, ...], bool]] = []
    for key in rule.keys:
        passes.append((key, False))
        if rule.swap_map and any(f in rule.swap_map for f in key):
            passes.append((key, True))

    matched: Dict[int, Tuple[int, bool]] = {}
    claimed: set = set()
    unmatched_a = list(range(len(section_a.rows)))

    for key_fields, reversed_pass in passes:
        if not unmatched_a:
            break
        index = _Index(section_b, key_fields)
        still: List[int] = []
        for ai in unmatched_a:
            key = _key_of(
                section_a,
                section_a.rows[ai],
                key_fields,
                rule.swap_map if reversed_pass else None,
            )
            bi = index.take(key, claimed)
            if bi is None:
                still.append(ai)
            else:
                matched[ai] = (bi, reversed_pass)
                claimed.add(bi)
        unmatched_a = still

    compared = [u for u in range(result.width) if u not in rule.ignore]

    for ai in range(len(section_a.rows)):
        row_a = section_a.rows[ai]
        key = _key_of(section_a, row_a, rule.keys[0])
        if ai not in matched:
            result.pairs.append(Pair(REMOVE, key, list(row_a), None))
            continue
        bi, reversed_pass = matched[ai]
        row_b = section_b.rows[bi]
        diffs = []
        for u in compared:
            # A reversed row has its paired fields the other way round, so
            # compare file 1 field 0 against file 2 field 3, not 0 against 0.
            ub = rule.swap_map.get(u, u) if reversed_pass else u
            value_a = section_a.field(row_a, u)
            value_b = section_b.field(row_b, ub)
            if value_a != value_b:
                diffs.append((u, ub, value_a, value_b))
        result.pairs.append(
            Pair(MODIFY if diffs else OK, key, list(row_a), list(row_b), diffs, reversed_pass)
        )

    for bi in range(len(section_b.rows)):
        if bi in claimed:
            continue
        row_b = section_b.rows[bi]
        result.pairs.append(Pair(ADD, _key_of(section_b, row_b, rule.keys[0]), None, list(row_b)))

    return result


def compare(dsi_a: Dsi, dsi_b: Dsi, only: Optional[Iterable[str]] = None) -> List[SectionResult]:
    """Diff every section present in either file, in file-1 order."""
    wanted = set(only) if only else None
    names: List[str] = []
    for section in dsi_a.sections:
        if section.name not in names:
            names.append(section.name)
    for section in dsi_b.sections:
        if section.name not in names:
            names.append(section.name)

    results = []
    for name in names:
        if name in SKIP_SECTIONS:
            continue
        if wanted is not None and name not in wanted:
            continue
        section_a = dsi_a.section(name)
        section_b = dsi_b.section(name)
        if (section_a is None or not section_a.rows) and (section_b is None or not section_b.rows):
            continue
        results.append(compare_section(name, section_a, section_b))
    return results


def field_header(section_name: str, index: int) -> str:
    labels = views.field_labels(section_name)
    label = labels.get(index)
    return "f{} {}".format(index, label) if label else "f{}".format(index)
