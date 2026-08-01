# Layout

`layout/` ships five documented python-docx gaps. None of them
duplicate functionality python-docx already exposes (orientation,
margins, page size, per-section header / footer, `add_section`).

For the calls, see the [layout guide](../guides/layout.md).

**`set_columns(section, num, *, space, separator, widths)`** in
`layout/columns.py` emits `<w:cols w:num=... w:space=... w:sep=...>`
into the section's `sectPr`. Idempotent — replaces any existing
`<w:cols>`. With `widths` supplied, it emits per-column `<w:col>`
children with `w:equalWidth="0"` so Word reads widths from the children
rather than the parent `w:space`.

**`insert_section_break(paragraph, *, start_type)`** in
`layout/breaks.py` handles the case `Document.add_section` does not —
inserting a break mid-document. The algorithm clones the trailing
body-level `<w:sectPr>` (the document's "sentinel"), sets `<w:type>`
on the clone to the requested start kind, and calls python-docx's
`CT_P.set_sectPr(clone)` to embed it in the chosen paragraph's `pPr`.
The new section inherits all properties (page size, margins, header /
footer references) from the sentinel; both sections render with the
same headers and footers unless the caller mutates the returned
`Section` proxy.

**`enable_distinct_even_odd_headers(doc)`** in `layout/settings.py`
writes `<w:evenAndOddHeaders/>` into `settings.xml` via the
[schema-strict insertion pattern](schema-order.md). This flag is
constantly confused with two other things: the per-section `<w:titlePg>`
(controls whether *first* page has a distinct header/footer, exposed by
python-docx as `Section.different_first_page_header_footer`), and the
per-section header/footer reference types (`w:headerReference w:type="even"`,
which Word reads *because* the doc-level flag is set). All three are
required for a real even-page-distinct workflow. `disable_…` removes
the doc-level element; both functions are idempotent.

**`set_line_numbering(section, *, count_by, restart, start, distance)`**
in `layout/line_numbering.py` emits `<w:lnNumType>` into the section's
`sectPr` — Word's mechanism for the marginal line numbers that legal
and contract documents require. Schema-strict via
`core.insert_before_first_anchor`; the element lands in its
ECMA-376 17.6.17 slot regardless of which other `sectPr` children
exist. `restart` is the only argument that validates eagerly (one of
`"newPage"` / `"newSection"` / `"continuous"`); `count_by` and `start`
must be ≥ 1. Idempotent.

**`set_page_borders(section, *, top, bottom, left, right)`** in
`layout/borders.py` emits `<w:pgBorders>` from a `Border` dataclass
per side (`style`, `size` in eighths of a point, `color`, `space` in
twips). Sides set to `None` are omitted from the emitted XML; passing
all four as `None` removes the element rather than emitting an empty
container. Schema-strict, idempotent.

The same `Border` shape drives [table and cell borders](tables.md), with
one caveat about `space` documented there.
