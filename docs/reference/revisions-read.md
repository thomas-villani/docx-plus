# `docx_plus.revisions.read`

Enumerate every tracked change in a document — run-level insertions and
deletions, move source/destination wrappers, run- and paragraph-property
changes, and paragraph-mark insertions/deletions — each paired with its id,
author, timestamp, type, and affected text. Insertion text is read from
`<w:t>`, deletion text from `<w:delText>`.

::: docx_plus.revisions.read
    options:
      members:
        - read_revisions
        - TrackedChange
        - RevisionType
