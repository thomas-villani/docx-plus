# `docx_plus.fields.simple`

Insert OOXML complex fields (PAGE / NUMPAGES / SECTIONPAGES, DATE /
CREATEDATE, plus a generic passthrough for everything else). Each
function emits the canonical 5-run sequence
(begin / instrText / separate / result-text / end) with
`xml:space="preserve"` on the instruction and the cached result so
Word does not collapse field-syntax whitespace.

Architecture walkthrough: [`ARCHITECTURE.md` §7](../ARCHITECTURE.md#7-fields-and-protection).

::: docx_plus.fields.simple
    options:
      members:
        - add_page_number_field
        - add_date_field
        - add_field
        - PageFieldName
