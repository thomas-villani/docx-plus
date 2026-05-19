# `docx_plus.bookmarks.crossref`

Cross-references to bookmarks via `REF` / `PAGEREF` complex fields.
Both are built on top of the same `core.build_complex_field` plumbing
that `fields/simple.py` uses for page numbers and dates. Pass
`kind="text"` for `REF` (resolves to the bookmark's text content) or
`kind="page"` for `PAGEREF` (resolves to the page number). The `\h`
flag is appended by default so Word renders the cross-reference as a
clickable link to the bookmark.

Pair calls with `docx_plus.fields.mark_fields_dirty` so Word
recalculates the cached results on first open.

::: docx_plus.bookmarks.crossref
    options:
      members:
        - add_cross_reference
        - CrossReferenceKind
