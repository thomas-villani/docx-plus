# `docx_plus.core.parts`

Package part / relationship plumbing for separate OOXML parts. v0.2
capabilities (`comments/`, `notes/`) live in their own parts under
`/word/comments.xml`, `/word/footnotes.xml`, and `/word/endnotes.xml`,
each of which may be absent from a fresh document. `get_or_create_part`
is the single entry point that returns the part and its parsed XML root,
fabricating both if missing.

Internal `XmlPart` subclasses for footnote and endnote content types are
registered with `PartFactory.part_type_for` at import time so existing
documents round-trip with parsed `.element` rather than raw blobs.

Architecture walkthrough: [`ARCHITECTURE.md` §7.5](../ARCHITECTURE.md#75-separate-ooxml-parts).

::: docx_plus.core.parts
    options:
      members:
        - get_or_create_part
        - PartSpec
        - COMMENTS_SPEC
        - FOOTNOTES_SPEC
        - ENDNOTES_SPEC
