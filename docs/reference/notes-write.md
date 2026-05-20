# `docx_plus.notes.write`

Insert footnotes and endnotes. Both share the same shape: a reference
marker run in the body, plus a content entry in the corresponding
separate part (`word/footnotes.xml` or `word/endnotes.xml`, created on
first use via `core.get_or_create_part`). `edit_footnote` / `edit_endnote`
replace the body text of an existing note in place; reserved ids (`-1`
separator, `0` continuation-separator) are not editable.

Architecture walkthrough: [`ARCHITECTURE.md` §7.9](../ARCHITECTURE.md#79-footnotes-and-endnotes).

::: docx_plus.notes.write
    options:
      members:
        - add_footnote
        - add_endnote
        - edit_footnote
        - edit_endnote
        - FootnoteRef
        - EndnoteRef
        - NoteNotFoundError
