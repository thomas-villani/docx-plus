# `docx_plus.layout.breaks`

Insert a section break at an arbitrary paragraph. python-docx's
`Document.add_section` only appends a new section at the end of the
document; this module handles the mid-document case by cloning the
trailing `<w:sectPr>` into the chosen paragraph's `pPr` and setting
`<w:type>` to the requested start kind.

::: docx_plus.layout.breaks
    options:
      members:
        - insert_section_break
        - SectionStartType
