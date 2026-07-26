# `docx_plus.core.borders`

The `CT_Border` shape, shared by every border edge in the format.

ECMA-376 uses one complex type for the four sides of `<w:pgBorders>`
(17.6.10), of `<w:tblBorders>` (17.4.39), and of `<w:tcBorders>`
(17.4.67), plus the inside and diagonal edges tables add — all with the
same four attributes. `Border` shipped from `docx_plus.layout` in v0.2
and moved here in v0.5 when table borders became a second consumer;
`docx_plus.layout.Border` still resolves to this class.

One default is worth knowing: `space` is `24` points, which is what
Word's UI emits for a *page* border. Word writes `w:space="0"` on table
and cell borders, so table writers override it rather than inheriting.

::: docx_plus.core.borders
    options:
      members:
        - Border
        - border_attrs
