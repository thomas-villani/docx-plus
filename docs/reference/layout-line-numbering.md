# `docx_plus.layout.line_numbering`

Section line numbering via `<w:lnNumType>`. python-docx does not
abstract the marginal-line-number element that legal and contract
documents commonly want. `set_line_numbering` emits the element with
schema-strict child ordering in `<w:sectPr>` and supports the four
ECMA-376 17.6.8 attributes: `countBy`, `restart`, `start`, and
`distance`.

Architecture walkthrough: [Layout](../concepts/layout.md).

::: docx_plus.layout.line_numbering
    options:
      members:
        - set_line_numbering
        - LineNumberRestart
