#!/usr/bin/env python3
"""Lossless reader/writer for Capital HarnessXC / ModularXC DSI export files.

A DSI export is flat text. Three characters, declared in the file's own header
banner, control its structure::

    ! * Comment Marker    =  !
    ! * Section Separator =  %
    ! * Field Delimiter   =  :

Lines beginning with the section separator open a section. Lines beginning with
the comment marker are one of four different things, told apart by what follows
the marker:

    ``! ****``, ``! * Created by:``   banner and export settings -> comment
    ``!``                             blank separator            -> comment
    ``!None``, ``! None``             section above has no rows  -> comment
    ``! terminals``                   a *section header*         -> section

That last case is not a quirk to be tolerated, it is load-bearing. The exporter
writes some section headers with the comment marker instead of the section
separator, and real data follows them::

    % Harness main node components      <- separator
      ...170 rows...
    ! terminals                         <- comment marker, but a header
      ...1216 rows...
    ! seals
    ! plug
    ! clips
    ! grommets
    ! other components
    ! Harness extra node components

Split on ``%`` alone and those seven blocks silently fuse into one 1707-row
section, and the first section of the file (``! Harness name information``)
disappears entirely because it sits above the first ``%``.

Round-trip guarantee: for a file with uniform line endings,

    Dsi.read(path).to_bytes() == open(path, 'rb').read()

Nothing is stripped, padded, reordered, or dropped -- including blank
separators, banner text, ``None`` markers and trailing empty fields -- so this
module is safe as the front end of a tool that rewrites DSI files and hands
them back to manufacturing.

Standard library only. Python 3.8+.
"""

from __future__ import annotations

import re
import sys
from typing import Dict, Iterator, List, Optional, Sequence, Tuple

__all__ = ["Dsi", "Section", "DsiError"]

DEFAULT_COMMENT = "!"
DEFAULT_SEPARATOR = "%"
DEFAULT_DELIMITER = ":"

# The banner declares the three control characters. Read them rather than
# assuming, but fall back to the documented defaults when absent.
_DECLARATIONS = (
    ("comment", re.compile(r"Comment\s+Marker\s*=\s*(\S)")),
    ("separator", re.compile(r"Section\s+Separator\s*=\s*(\S)")),
    ("delimiter", re.compile(r"Field\s+Delimiter\s*=\s*(\S)")),
)

_LINE_SPLIT = re.compile(r"\r\n|\r|\n")

# A comment line is a *header* unless its payload is empty, starts with '*'
# (banner / settings), or is the empty-section marker.
_NONE_PAYLOAD = re.compile(r"^none$", re.IGNORECASE)


class DsiError(Exception):
    """Raised when a DSI file cannot be parsed."""


class Section:
    """One section of a DSI file: a header line plus its data rows.

    Attributes:
        name: header text with the marker and surrounding space removed,
            e.g. ``'Harness branch configuration'``.
        marker: the character that introduced the header -- the section
            separator (``%``) for ordinary sections, the comment marker (``!``)
            for the sub-blocks and for ``Harness name information``.
        rows: one ``list[str]`` per data line, fields verbatim and ragged.
            Row lengths genuinely vary within a section; nothing is padded.
        line_no: 0-based line index of the header in the source file.
        parent: name of the enclosing separator-marked section for a
            comment-marked sub-block, else ``None``.
        lead: raw comment lines that sat between the previous section's last
            row and this header. Preserved only so round-trips stay exact.
        notes: ``(row_index, raw_line)`` for comment lines inside the body;
            ``row_index`` is how many rows preceded the comment. This is where
            ``!None`` lands for an empty section.
    """

    __slots__ = (
        "name",
        "marker",
        "rows",
        "line_no",
        "parent",
        "lead",
        "notes",
        "comment_marker",
        "_row_lines",
    )

    def __init__(
        self,
        name: str,
        marker: str,
        line_no: int,
        parent: Optional[str] = None,
        lead: Optional[List[str]] = None,
        comment_marker: str = DEFAULT_COMMENT,
    ) -> None:
        self.name = name
        self.marker = marker
        self.line_no = line_no
        self.parent = parent
        self.lead: List[str] = lead or []
        self.comment_marker = comment_marker
        self.rows: List[List[str]] = []
        self.notes: List[Tuple[int, str]] = []
        self._row_lines: List[int] = []

    # -- shape ---------------------------------------------------------------

    def __len__(self) -> int:
        return len(self.rows)

    def __iter__(self) -> Iterator[List[str]]:
        return iter(self.rows)

    def __getitem__(self, index: int) -> List[str]:
        return self.rows[index]

    def __repr__(self) -> str:
        return "<Section {!r} marker={!r} rows={} fields={}>".format(
            self.name, self.marker, len(self.rows), self.field_range()
        )

    @property
    def width(self) -> int:
        """Field count of the widest row, 0 when the section is empty."""
        return max((len(r) for r in self.rows), default=0)

    def field_range(self) -> Tuple[int, int]:
        """``(min, max)`` field count across rows, ``(0, 0)`` when empty."""
        if not self.rows:
            return (0, 0)
        lengths = [len(r) for r in self.rows]
        return (min(lengths), max(lengths))

    @property
    def is_ragged(self) -> bool:
        lo, hi = self.field_range()
        return lo != hi

    @property
    def declared_empty(self) -> bool:
        """True when the exporter wrote an explicit ``None`` marker.

        A section can be empty two ways: an explicit marker, or a header with
        simply nothing under it. Worth telling apart -- the first says the
        exporter considered the section and found no data, the second may mean
        the export was truncated.
        """
        if self.rows:
            return False
        marker = self.comment_marker
        for _, raw in self.notes:
            payload = raw[len(marker):] if raw.startswith(marker) else raw
            if _NONE_PAYLOAD.match(payload.strip()):
                return True
        return False

    def source_line(self, row_index: int) -> int:
        """0-based line number in the source file for ``rows[row_index]``."""
        return self._row_lines[row_index]

    # -- field access --------------------------------------------------------

    def field(self, row: Sequence[str], index: int, default: str = "") -> str:
        """``row[index]`` when present, else ``default``.

        Rows are ragged, so positional access needs a floor. Returning a
        default here rather than raising mirrors how the source data behaves:
        a trailing field that was never written is semantically empty.
        """
        if 0 <= index < len(row):
            return row[index]
        return default

    def key(self, row: Sequence[str], indices: Sequence[int]) -> Tuple[str, ...]:
        """Tuple of the given field positions -- the match key for diffing."""
        return tuple(self.field(row, i) for i in indices)

    def column(self, index: int) -> List[str]:
        """Field ``index`` from every row, empty string where absent."""
        return [self.field(r, index) for r in self.rows]


class Dsi:
    """A parsed DSI file.

    Sections keep source order. Look one up by exact name with ``section()``
    or ``dsi['Harness wire specification']``.
    """

    def __init__(
        self,
        sections: List[Section],
        preamble: List[str],
        comment: str = DEFAULT_COMMENT,
        separator: str = DEFAULT_SEPARATOR,
        delimiter: str = DEFAULT_DELIMITER,
        newline: str = "\r\n",
        final_newline: bool = True,
        mixed_newlines: bool = False,
        encoding: str = "utf-8",
        path: Optional[str] = None,
    ) -> None:
        self.sections = sections
        self.preamble = preamble
        self.comment = comment
        self.separator = separator
        self.delimiter = delimiter
        self.newline = newline
        self.final_newline = final_newline
        self.mixed_newlines = mixed_newlines
        self.encoding = encoding
        self.path = path

    # -- construction --------------------------------------------------------

    @classmethod
    def read(cls, path: str, encoding: str = "utf-8") -> "Dsi":
        with open(path, "rb") as handle:
            raw = handle.read()
        if raw.startswith(b"\xef\xbb\xbf"):
            raw = raw[3:]
            encoding = "utf-8-sig"
        return cls.parse(raw.decode(encoding), path=path, encoding=encoding)

    @classmethod
    def parse(cls, text: str, path: Optional[str] = None, encoding: str = "utf-8") -> "Dsi":
        endings = _LINE_SPLIT.findall(text)
        newline = endings[0] if endings else "\r\n"
        mixed = len(set(endings)) > 1
        lines = _LINE_SPLIT.split(text)
        final_newline = bool(lines) and lines[-1] == ""
        if final_newline:
            lines.pop()

        comment, separator, delimiter = cls._detect_markers(lines)

        sections: List[Section] = []
        preamble: List[str] = []
        pending: List[str] = []
        current: Optional[Section] = None
        last_top_level: Optional[str] = None

        for line_no, line in enumerate(lines):
            if line.startswith(separator):
                name = line[len(separator):].strip()
                current = Section(name, separator, line_no, None, pending, comment)
                last_top_level = name
                pending = []
                sections.append(current)
            elif line.startswith(comment):
                payload = line[len(comment):].strip()
                if payload and not payload.startswith("*") and not _NONE_PAYLOAD.match(payload):
                    # A header wearing the comment marker.
                    current = Section(payload, comment, line_no, last_top_level, pending, comment)
                    pending = []
                    sections.append(current)
                elif current is None:
                    preamble.append(line)
                elif _NONE_PAYLOAD.match(payload):
                    # The marker describes the section above it, so keep it
                    # there rather than letting it drift into the next lead.
                    for raw in pending:
                        current.notes.append((len(current.rows), raw))
                    pending = []
                    current.notes.append((len(current.rows), line))
                else:
                    pending.append(line)
            else:
                if current is None:
                    raise DsiError(
                        "data on line {} before any section header: {!r}".format(line_no + 1, line[:60])
                    )
                if pending:
                    # Comments that turned out to sit *inside* a body.
                    for raw in pending:
                        current.notes.append((len(current.rows), raw))
                    pending = []
                current.rows.append(line.split(delimiter))
                current._row_lines.append(line_no)

        # Comments trailing the final section belong to it.
        if pending:
            if current is None:
                preamble.extend(pending)
            else:
                for raw in pending:
                    current.notes.append((len(current.rows), raw))

        return cls(
            sections=sections,
            preamble=preamble,
            comment=comment,
            separator=separator,
            delimiter=delimiter,
            newline=newline,
            final_newline=final_newline,
            mixed_newlines=mixed,
            encoding=encoding,
            path=path,
        )

    @staticmethod
    def _detect_markers(lines: Sequence[str]) -> Tuple[str, str, str]:
        found = {"comment": DEFAULT_COMMENT, "separator": DEFAULT_SEPARATOR, "delimiter": DEFAULT_DELIMITER}
        # The banner is at the top; 40 lines is generous and bounds the scan.
        for line in lines[:40]:
            for key, pattern in _DECLARATIONS:
                match = pattern.search(line)
                if match:
                    found[key] = match.group(1)
        if len({found["comment"], found["separator"], found["delimiter"]}) != 3:
            raise DsiError("control characters are not distinct: {!r}".format(found))
        return found["comment"], found["separator"], found["delimiter"]

    # -- lookup --------------------------------------------------------------

    def __iter__(self) -> Iterator[Section]:
        return iter(self.sections)

    def __len__(self) -> int:
        return len(self.sections)

    def __getitem__(self, name: str) -> Section:
        section = self.section(name)
        if section is None:
            raise KeyError(name)
        return section

    def __contains__(self, name: str) -> bool:
        return self.section(name) is not None

    def __repr__(self) -> str:
        return "<Dsi {!r} sections={} rows={}>".format(self.path, len(self.sections), self.row_count)

    def section(self, name: str) -> Optional[Section]:
        """First section with this exact name, or ``None``.

        Name matching is exact by design. The VBA this replaces keys on the
        full header string, and the two engine families differ only in those
        strings (``! terminals`` vs ``! Harness terminals``), so loose matching
        would blur the one distinction that identifies the source tool.
        """
        for section in self.sections:
            if section.name == name:
                return section
        return None

    def find(self, *names: str) -> Optional[Section]:
        """First section matching any of ``names``, in the order given.

        Lets a caller accept either engine family's spelling::

            dsi.find('terminals', 'Harness terminals')
        """
        for name in names:
            section = self.section(name)
            if section is not None:
                return section
        return None

    def children(self, name: str) -> List[Section]:
        """Comment-marked sub-blocks nested under a separator-marked section."""
        return [s for s in self.sections if s.parent == name and s.marker == self.comment]

    @property
    def names(self) -> List[str]:
        return [s.name for s in self.sections]

    @property
    def row_count(self) -> int:
        return sum(len(s.rows) for s in self.sections)

    def duplicate_names(self) -> List[str]:
        """Names appearing on more than one section, which make lookup ambiguous."""
        seen: Dict[str, int] = {}
        for section in self.sections:
            seen[section.name] = seen.get(section.name, 0) + 1
        return sorted(n for n, c in seen.items() if c > 1)

    # -- serialisation -------------------------------------------------------

    def to_lines(self) -> List[str]:
        out: List[str] = list(self.preamble)
        for section in self.sections:
            out.extend(section.lead)
            out.append(section.marker + " " + section.name if section.name else section.marker)
            notes = {}
            for at, raw in section.notes:
                notes.setdefault(at, []).append(raw)
            for index, row in enumerate(section.rows):
                out.extend(notes.pop(index, ()))
                out.append(self.delimiter.join(row))
            for at in sorted(notes):
                out.extend(notes[at])
        return out

    def to_text(self) -> str:
        text = self.newline.join(self.to_lines())
        if self.final_newline:
            text += self.newline
        return text

    def to_bytes(self) -> bytes:
        return self.to_text().encode(self.encoding)

    def write(self, path: str) -> None:
        with open(path, "wb") as handle:
            handle.write(self.to_bytes())

    def verify_round_trip(self, path: Optional[str] = None) -> bool:
        """True when re-serialising reproduces the source file byte for byte."""
        target = path or self.path
        if target is None:
            raise DsiError("no path to verify against")
        with open(target, "rb") as handle:
            original = handle.read()
        if original.startswith(b"\xef\xbb\xbf"):
            original = original[3:]
        return original == self.to_bytes()


# -- command line -----------------------------------------------------------


def _census(dsi: Dsi) -> str:
    rows = [("SECTION", "MK", "ROWS", "FIELDS", "PARENT")]
    for section in dsi.sections:
        lo, hi = section.field_range()
        if not section.rows:
            shape = "empty" if section.declared_empty else "-"
        elif lo == hi:
            shape = str(lo)
        else:
            shape = "{}-{}".format(lo, hi)
        rows.append(
            (section.name, section.marker, str(len(section.rows)), shape, section.parent or "")
        )
    widths = [max(len(r[i]) for r in rows) for i in range(5)]
    lines = []
    for index, row in enumerate(rows):
        lines.append("  ".join(cell.ljust(widths[i]) for i, cell in enumerate(row)).rstrip())
        if index == 0:
            lines.append("  ".join("-" * widths[i] for i in range(5)))
    return "\n".join(lines)


def main(argv: Sequence[str]) -> int:
    if len(argv) != 2:
        print("usage: dsi.py <file.dsi>", file=sys.stderr)
        print("       prints a section census and verifies the round trip", file=sys.stderr)
        return 2
    path = argv[1]
    try:
        dsi = Dsi.read(path)
    except (OSError, DsiError, UnicodeDecodeError) as exc:
        print("cannot read {}: {}".format(path, exc), file=sys.stderr)
        return 1

    print("file        {}".format(path))
    print("markers     comment={!r} separator={!r} delimiter={!r}".format(dsi.comment, dsi.separator, dsi.delimiter))
    print("newline     {!r}{}".format(dsi.newline, "  MIXED" if dsi.mixed_newlines else ""))
    print("sections    {}   rows {}".format(len(dsi.sections), dsi.row_count))
    duplicates = dsi.duplicate_names()
    if duplicates:
        print("duplicates  {}".format(", ".join(duplicates)))
    print()
    print(_census(dsi))
    print()
    exact = dsi.verify_round_trip()
    print("round trip  {}".format("byte-identical" if exact else "DIFFERS -- parser is lossy"))
    return 0 if exact else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
