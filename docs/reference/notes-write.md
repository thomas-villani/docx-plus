# `docx_plus.notes.write`

Insert footnotes and endnotes. Both share the same shape: a reference
marker run in the body, plus a content entry in the corresponding
separate part (`word/footnotes.xml` or `word/endnotes.xml`, created on
first use via `core.get_or_create_part`). Insert-only is sufficient for
v0.2 — in-place edits of existing notes are deferred to v0.3.

Architecture walkthrough: [`ARCHITECTURE.md` §7.9](../ARCHITECTURE.md#79-footnotes-and-endnotes).

::: docx_plus.notes.write
    options:
      members:
        - add_footnote
        - add_endnote
        - FootnoteRef
        - EndnoteRef
