# `docx_plus.layout.columns`

Multi-column page layout via `<w:cols>`. python-docx exposes
orientation, margins, page size, and headers/footers on
`docx.section.Section`, but does not abstract column layout. This
module's single `set_columns` helper fills the gap, supporting equal
columns, unequal columns via the `widths` argument, and the optional
vertical separator line (`w:sep="1"`).

Architecture walkthrough: [`ARCHITECTURE.md` §7.7](../ARCHITECTURE.md#77-layout).

::: docx_plus.layout.columns
    options:
      members:
        - set_columns
