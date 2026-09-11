# DSI Toolkit

Python tools for Capital HarnessXC / ModularXC DSI exports. **Standard library
only** — no `pip install`, runs on a stock Python 3.8+.

## Run it

```
python3 dsi_tool.py
```

It asks how many files you have:

- **One file** → asks which elements you want (connectors, splices, terminals,
  wires, branches, insulations, …) and writes a formatted workbook, one sheet
  per element.
- **Two files** → compares them and writes a difference report.

Non-interactive equivalents:

```
python3 dsi_tool.py view    HARNESS.dsi
python3 dsi_tool.py view    HARNESS.dsi --elements connectors,splices,wires
python3 dsi_tool.py compare OLD.dsi NEW.dsi
python3 dsi_tool.py compare OLD.dsi NEW.dsi --all-rows -o REPORT.xlsx
python3 dsi_tool.py --list-elements
```

## The comparison report

| Sheet | What it holds |
|---|---|
| `SYNTHESIS` | Per section: row counts, OK / MODIFY / ADD / REMOVE / REVERSED, **and the match key used** |
| `DIFFERENCES` | Every delta as one row: section, key, field, value in each file |
| one per section | Side by side, file 1 block then file 2 block, differing cells filled |
| `FILES` | Export banner of both files |

By default unchanged rows are left out of the section sheets. Pass `--all-rows`
to keep them.

`REVERSED` counts segments recorded A→B in one file and B→A in the other. They
are matched as the same segment rather than reported as one ADD plus one
REMOVE, and the direction flip is still listed on `DIFFERENCES`.

## Changing how rows are matched

Every section's match key lives in one table, `RULES` in `compare.py`:

```python
Rule(
    ["Harness branch configuration"],
    [(0, 3)],                            # match on the two end nodes
    swap=[(0, 3), (2, 5)],               # A->B and B->A are the same branch
    ignore={2, 5, 19, 20, 21, 22, 23},   # coordinates: don't report geometry
    note="...",                          # shown on the SYNTHESIS sheet
)
```

`SYNTHESIS` prints the key and ignored fields actually used for each section, so
you can check the result against the VBA output and adjust one line here rather
than hunting through sixty near-identical subroutines.

If you want geometry differences reported, delete the coordinate fields from
`ignore`.

## Modules

| File | Role |
|---|---|
| `dsi.py` | Parse a DSI into sections and rows. Round-trips byte for byte, so it is safe to use for rewriting DSI files. Run it directly for a section census. |
| `views.py` | Element definitions for single-file mode: which section, which rows, which fields under which heading |
| `compare.py` | The match-key table and the one comparison function |
| `xlsx.py` | Small xlsx writer with per-cell fills |
| `dsi_tool.py` | The interactive front end |
| `dsi_to_excel.py` | Earlier whole-file dump, kept as is |

## Notes on the format

Established by reading the sample export, not from documentation:

- **Section headers use both control characters.** Seven headers are written
  with the comment marker instead of the section separator — `! terminals`,
  `! seals`, `! plug`, `! clips`, `! grommets`, `! other components`,
  `! Harness extra node components` — and they carry real data. Splitting on
  `%` alone fuses them into one 1707-row section and drops
  `! Harness name information` entirely.
- **Main node components field 4** is the type discriminator:
  `CONNECTOR` / `SPLICE` / `IDC`. That is what separates connectors from
  splices.
- **Branch coordinates are 3D**, packed as `x1820.00y-905.00z0.00`.
- **Option expressions**: `+` is AND, `/` is OR, `!` is NOT, parentheses group.
  They live in wire specification field 1, branch insulations field 7, branch
  configuration field 7 and clips field 5.
- **Circuit option codes start at field 36.**
