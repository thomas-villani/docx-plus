# `docx_plus.layout.borders`

Page borders via `<w:pgBorders>` with a per-side `Border` dataclass.
Each side carries style (e.g. `"single"`, `"double"`), thickness in
eighths of a point, RGB hex color, and a `space` gap in points
(range 0-31). `set_page_borders` also takes `offset_from` (`"page"`
default, matching Word's UI, or `"text"`). All four sides default to
`None`; passing all-None removes the element rather than emitting an
empty container.

Architecture walkthrough: [`ARCHITECTURE.md` §7.7](../ARCHITECTURE.md#77-layout).

::: docx_plus.layout.borders
    options:
      members:
        - Border
        - set_page_borders
        - OffsetFrom
