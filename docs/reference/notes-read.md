# `docx_plus.notes.read`

Read footnotes and endnotes from a document. Each result is paired with
the paragraph index of its reference marker so callers can locate where
the note is referenced. Reserved entries (separator and continuation
separator, ids `-1` and `0`) are filtered out before results are
returned.

::: docx_plus.notes.read
    options:
      members:
        - read_footnotes
        - read_endnotes
        - NoteContent
