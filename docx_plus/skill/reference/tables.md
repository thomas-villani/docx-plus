# Tables — borders, shading, merging

python-docx models table *structure* well and table *appearance* not at
all. Rows, columns, cells, cell text, widths, and `_Cell.merge` all
work. But there is no `CT_Border`, `CT_TblBorders`, `CT_TcBorders`, or
`CT_Shd` class in the package and none of those tags is registered — so
ruling a table or shading a header row means raw OOXML.

Module: `docx_plus.tables`. Structure still comes from python-docx
(`doc.add_table`, `table.cell`, `cell.text`, `cell.width`); this module
only formats what is already there.

> **Not fields.** Nothing here needs `mark_fields_dirty`.

## Borders

```python
from docx_plus.core import Border
from docx_plus.tables import set_table_borders, set_cell_borders

table = doc.add_table(rows=4, cols=3)

hairline = Border(style="single", size=4, color="8EAADB")
set_table_borders(
    table,
    all_edges=Border(style="single", size=12, color="2F5496"),  # heavy box
    inside_h=hairline,                                          # light rules
    inside_v=hairline,
)

# A cell-level border overrides the table's for that cell:
set_cell_borders(table.cell(2, 2), top=Border(style="double", size=6, color="C00000"))
```

- `set_table_borders(table, *, all_edges=None, top=None, bottom=None,
  left=None, right=None, inside_h=None, inside_v=None)`
- `set_cell_borders(cell, *, all_edges=None, top=None, bottom=None,
  left=None, right=None, tl2br=None, tr2bl=None)` — `tl2br` / `tr2bl`
  are the diagonals. `all_edges` covers the four sides *only*; a
  diagonal is a "crossed-out cell" mark, never what "all borders" means.

Both are a **full replacement**, not a merge: an edge you do not name
ends up absent. Naming no edges removes the element entirely.

**`Border.size` is in eighths of a point** — `4` is 0.5 pt, `8` is 1 pt,
capped at 96.

> **`Border.space` is ignored for tables.** Both writers emit
> `w:space="0"`. The dataclass defaults it to `24`, a *page*-border
> value; Word's table UI cannot produce anything but `0`. Nothing to do
> — just don't expect the field to survive a round-trip.

## Shading

```python
from docx_plus.tables import Shading, set_table_shading, set_row_shading, set_cell_shading

set_row_shading(table.rows[0], Shading(fill="2F5496"))      # header band
set_cell_shading(table.cell(3, 2), Shading(fill="FFF2CC"))  # one call-out cell
set_table_shading(table, Shading(fill="F2F2F2"))            # default for the whole table
set_cell_shading(table.cell(3, 2), None)                    # remove
```

`Shading(fill="auto", pattern="clear", color="auto")` — `fill` is the
background and is what you almost always want alone. `pattern` is an
`ST_Shd` value drawn over it (`"pct25"`, `"thinHorzStripe"`, `"nil"`);
`color` is that pattern's foreground and does nothing while `pattern`
is `"clear"`. All three validate at construction — `fill="red"` or
`"#FF0000"` raises.

> **Rows have no shading element.** `CT_TrPr` has no `w:shd` child.
> Word implements "shade this row" by writing into every cell, and so
> does `set_row_shading`. It iterates `<w:tc>` elements rather than
> `Row.cells`, so a cell spanning several columns is written once.

Shading is a *cell* property; contrasting text on top of it is ordinary
python-docx run formatting:

```python
from docx.shared import RGBColor

# A fresh cell holds an empty paragraph and no runs, so add one rather
# than indexing into `.runs` — `paragraphs[0].runs[0]` is an IndexError
# until something has written text.
run = table.cell(0, 0).paragraphs[0].add_run("Header")
run.bold = True
run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
```

## Merging and unmerging

```python
from docx_plus.tables import merge_cells, unmerge_cell

banner = merge_cells(table.cell(0, 0), table.cell(0, 2))   # returns the top-left cell
banner.text = "Fiscal Year 2026"

unmerge_cell(banner)          # the inverse — python-docx has none
```

- `merge_cells(start, end)` — the two cells are diagonal corners of a
  rectangular region. Content from every cell moves into the top-left
  one, which is what is returned and is not necessarily `start`. A
  non-rectangular (L- or T-shaped) selection raises `InvalidMergeError`
  (`DocxPlusError` + `ValueError`).
- `unmerge_cell(cell)` — works from *any* cell in the region, including
  a vertical continuation. Content stays in the anchor; the restored
  cells are empty, matching Word's "Split Cells". Widths divide evenly,
  because the originals were summed away by the merge. Idempotent on an
  unmerged cell.

## Legacy `w:hMerge` tables

OOXML has **two** horizontal-merge encodings and python-docx models only
one:

| Encoding | Shape | python-docx |
|---|---|---|
| `w:gridSpan` | One `<w:tc>` widened over N columns | Understood |
| `w:hMerge` | One `<w:tc>` per column, followers marked continuation | Ignored |

Word renders both identically. But on an `hMerge` table `table.cell()`
hands back cells that look separate and are not, and `row.cells` reports
a column count Word never draws. If a table you did not author behaves
strangely, this is a likely cause:

```python
from docx_plus.tables import normalize_horizontal_merges

converted = normalize_horizontal_merges(table)          # -> number of regions
converted = normalize_horizontal_merges(table, discard_content=True)
```

Returns `0` on any table Word wrote recently. It refuses by default to
drop text held in a continuation cell — that text is invisible in Word,
so keeping it would surface hidden content and dropping it silently
would lose data. Pass `discard_content=True` to choose.

## Reading it back

```python
from docx_plus.tables import read_table_formatting

fmt = read_table_formatting(table)
print(fmt.style, sorted(fmt.borders), fmt.shading)
for cell in fmt.cells:
    print(cell.row, cell.column, cell.grid_span, cell.vertical_merge, cell.shading)
```

> **Direct formatting only.** The table-style cascade (table style →
> `w:tblStylePr` conditional branch → direct `w:tcPr`) is **not**
> resolved. A `Table Grid` table reads back with `borders == {}` — true
> of its XML, not of its appearance. Do not use this to answer "does
> this table have borders?"; use it to answer "what did someone write
> directly on it?".

`CellFormatting.column` is a **grid offset**, not an index into
`row.cells` — cells right of a merged one are offset by the span. One
entry per `<w:tc>`, so a merged cell appears once.

## End-to-end

```python
from docx import Document
from docx_plus.core import Border
from docx_plus.tables import Shading, merge_cells, set_row_shading, set_table_borders

doc = Document()
table = doc.add_table(rows=4, cols=3)

banner = merge_cells(table.cell(0, 0), table.cell(0, 2))
banner.text = "Fiscal Year 2026"
set_row_shading(table.rows[0], Shading(fill="2F5496"))

hairline = Border(style="single", size=4, color="8EAADB")
set_table_borders(table, all_edges=Border(style="single", size=12,
                                          color="2F5496"),
                  inside_h=hairline, inside_v=hairline)
doc.save("budget.docx")
```

See also: `docx_plus/examples/table_formatting.py`. Page borders are a
different module — `docx_plus.layout.set_page_borders`, in
`reference/layout.md`.
