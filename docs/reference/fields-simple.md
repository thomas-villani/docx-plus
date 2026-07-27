# `docx_plus.fields.simple`

Insert OOXML complex fields (PAGE / NUMPAGES / SECTIONPAGES, DATE /
CREATEDATE, plus a generic passthrough for everything else). Each
function emits the canonical 5-run sequence
(begin / instrText / separate / result-text / end) with
`xml:space="preserve"` on the instruction and the cached result so
Word does not collapse field-syntax whitespace.

`add_style_reference` (v0.5) emits `STYLEREF`, the one cross-reference
that needs no bookmark: it resolves to the text of the nearest paragraph
carrying a given style, re-evaluated per page. That is what makes a
running header show the current chapter — verified against Word 2016,
where the same field renders `Chapter: Architecture` on page 1 and
`Chapter: Operations` on page 2.

Note its `style` argument takes the style *name* as Word shows it
(`"Heading 1"`, with the space), not the `w:styleId` — unlike most of
this library, because that is what the field instruction accepts. An
`int` is an outline level instead.

Architecture walkthrough: [`ARCHITECTURE.md` §7](../ARCHITECTURE.md#7-fields-and-protection).

::: docx_plus.fields.simple
    options:
      members:
        - add_page_number_field
        - add_date_field
        - add_style_reference
        - add_field
        - PageFieldName
        - StyleRefNumber
