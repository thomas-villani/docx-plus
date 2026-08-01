# Table formatting

`tables/` (v0.5) covers the half of tables python-docx leaves out.

python-docx models table **structure** well: rows, columns, cells, cell
text, widths, and a working `_Cell.merge`. It models table
**appearance** not at all. There is no `CT_Border`, no `CT_TblBorders`,
no `CT_TcBorders`, and no `CT_Shd` class in the package, and none of
those tags is registered — so a border or fill written by hand
round-trips as an anonymous `lxml` element.

This is deliberately *not* part of [`layout/`](layout.md): every helper
there takes a `Section` or `Document` and its docstring scopes it to page
layout.

For the calls, see the [tables guide](../guides/tables.md).

## Borders and shading

Structurally these are `set_page_borders` again — the same
`CT_Border` shape from `core/borders.py`, the same schema-ordered
insertion, the same replace-or-remove idempotence. Tables add the two
inside edges (17.4.39); cells add the two diagonals (17.4.67).

The one non-obvious point is **`w:space`**. `Border.space` defaults to
`24`, a *page* value: what Word emits for "Whole document, Box, Default
settings". Word's UI cannot produce a non-zero space on a table border
at all and always writes `0`. Reusing the dataclass default blindly
would put a third of an inch between every table edge and its text, so
both writers pin the attribute to `0` and say so.

Row shading needs its own note: **`CT_TrPr` has no `w:shd` child.**
There is no row-level shading in the format. Word implements "shade
this row" by writing the same `<w:shd>` into every cell, and so does
`set_row_shading`. It iterates the row's `<w:tc>` elements rather than
`Row.cells`, so a cell spanning several grid columns is visited once
rather than once per column it covers.

## The two horizontal-merge encodings

OOXML can express a horizontal merge two ways:

- **`w:gridSpan`** (17.4.17) — one `<w:tc>` widened over several grid
  columns. This is what Word writes today and the only form python-docx
  understands.
- **`w:hMerge`** (17.4.22) — one `<w:tc>` per column, followers marked
  as continuations. Older Word versions and several converters emit
  this.

Word renders them identically — verified against Word 2016, where a
converted file rasterises byte-for-byte the same as its original. But
python-docx's grid model ignores `hMerge` entirely, so `Table.cell`
hands back cells that look separate and are not. Word's own COM object
model shares the blind spot: it reported six cells for the `hMerge`
fixture and five after conversion, while laying both out the same way.

`normalize_horizontal_merges` rewrites the second form as the first.
It refuses by default to drop text held in a continuation cell —
invisible in Word, so keeping it would make hidden content appear and
discarding it silently would lose data.

Note that "has content" cannot be "has a `<w:r>`": every cell holds at
least one `<w:p>`, and `cell.text = ""` leaves an empty run behind, so
that test calls every ordinary cell occupied. The check looks for
non-blank `<w:t>` text or an embedded object.

## Unmerging

`_Cell.merge` is fully implemented in python-docx and is **not**
re-implemented here — `merge_cells` only translates `InvalidSpanError`
into a `DocxPlusError` subclass per the [error
hierarchy](invariants.md#error-hierarchy). The inverse is what is
missing: nothing in python-docx removes a `w:gridSpan` or a `w:vMerge`,
so a merge is one-way.

`unmerge_cell` resolves the region from any cell in it, including a
vertical continuation, then walks the vertical run *before* mutating
anything — splitting a cell horizontally shifts the grid offsets of its
right-hand neighbours, so the lookups have to happen first. Widths are
divided evenly, because the individual widths were summed away when the
merge happened and cannot be recovered.

## Not covered

The **cell-formatting cascade** (table style → `<w:tblStylePr>`
conditional branch → direct `<w:tcPr>`). `read_table_formatting`
reports direct formatting only, so a `Table Grid` table reads back with
no borders — true of its XML, not of its appearance. The [cascade
resolver](cascade.md) covers paragraphs and runs and scopes this out in
the same terms; it is a larger workstream than every writer in this
package put together.
