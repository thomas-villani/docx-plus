# `docx_plus.tables.borders`

Table and cell borders. python-docx has no element class for any of
this — no `CT_Border`, no `CT_TblBorders`, no `CT_TcBorders`, and none
of those tags is registered — so ruling a table has meant writing OOXML
by hand.

Both writers are a **full replacement**, not a merge: any existing
container is discarded first, so an edge you do not name ends up
absent. Naming no edges at all removes the element rather than leaving
an empty container behind.

!!! warning "`Border.space` is ignored here"
    Both writers emit `w:space="0"`, which is what Word does and the
    only value its UI can produce for a table. The
    [`Border`](core-borders.md) dataclass defaults `space` to `24` — a
    *page*-border value — which would otherwise put a third of an inch
    between every table edge and its text.

Architecture walkthrough:
[Table formatting](../concepts/tables.md).

::: docx_plus.tables.borders
    options:
      members:
        - set_table_borders
        - set_cell_borders
