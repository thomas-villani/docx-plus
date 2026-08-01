# `docx_plus.tables.read`

Reads back the borders, shading, and merge state actually written on a
table, its rows, and its cells.

!!! warning "Direct formatting only"
    This reports what is present on the `<w:tblPr>` and `<w:tcPr>`
    elements themselves. It does **not** resolve the cell-formatting
    cascade (table style → `<w:tblStylePr>` conditional branch → direct
    properties), so a table whose ruling comes entirely from a style
    such as `Table Grid` reads back with no borders at all — the truth
    about its XML, not about its appearance.

    That resolver is a considerably larger piece of work than every
    writer in this package put together.
    [`docx_plus.styles.inspect`](styles-inspect.md) resolves the
    paragraph and run cascade but scopes this one out in the same
    terms.

`CellFormatting.column` is a **grid offset**, not an index into
`Row.cells`: cells to the right of a merged one are offset by the span.
One entry is produced per `<w:tc>` element, so a merged cell appears
once rather than once per grid column it covers.

Note that `space` on a returned `Border` is always `0` — see
[`tables.borders`](tables-borders.md) for why the writers pin it.

Architecture walkthrough:
[Table formatting](../concepts/tables.md).

::: docx_plus.tables.read
    options:
      members:
        - read_table_formatting
        - TableFormatting
        - CellFormatting
