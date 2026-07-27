# `docx_plus.tables.shading`

Cell, row, and table background fills. python-docx has no `CT_Shd`
class and does not register the `w:shd` tag, so the single most common
thing anyone wants from a table beyond its text is unreachable.

ECMA-376 models shading as three attributes rather than one colour: a
`w:fill` (the background), a `w:val` pattern drawn over it, and a
`w:color` for that pattern's foreground. A solid fill — what nearly
everyone means — is `pattern="clear"` with only `fill` set, which is
what [`Shading`](#docx_plus.tables.shading.Shading) defaults to.

!!! note "Rows have no shading element"
    `CT_TrPr` has no `w:shd` child. Word implements "shade this row" by
    writing the same `<w:shd>` into every cell in it, and so does
    `set_row_shading`. It iterates the row's `<w:tc>` elements rather
    than `Row.cells`, so a cell spanning several grid columns is
    visited once rather than once per column.

Architecture walkthrough:
[`ARCHITECTURE.md` §7.14](../ARCHITECTURE.md#714-table-formatting).

::: docx_plus.tables.shading
    options:
      members:
        - Shading
        - set_table_shading
        - set_row_shading
        - set_cell_shading
        - shading_attrs
