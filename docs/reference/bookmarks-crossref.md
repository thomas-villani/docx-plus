# `docx_plus.bookmarks.crossref`

Cross-references to bookmarks via `REF` / `PAGEREF` complex fields.
Both are built on top of the same `core.build_complex_field` plumbing
that `fields/simple.py` uses for page numbers and dates. Pass
`kind="text"` for `REF` (resolves to the bookmark's text content) or
`kind="page"` for `PAGEREF` (resolves to the page number). The `\h`
flag is appended by default so Word renders the cross-reference as a
clickable link to the bookmark.

The switches carry more weight than they look. Verified against Word
2016, the same bookmark yields:

| Call | Resolves to |
|---|---|
| `add_cross_reference(p, bookmark=fig)` | `Figure 1` |
| `..., kind="page"` | `1` |
| `..., position=True` | `above` |
| `..., number="relative"` | the target's paragraph number, e.g. `2.3` |

The first row is only useful if something bookmarked the caption —
a `REF` **cannot point at a `SEQ` field**. See `bookmark_name` on
[`add_caption`](publishing-captions.md).

For a cross-reference that needs no bookmark at all, see
[`add_style_reference`](fields-simple.md) (`STYLEREF`).

Pair calls with `docx_plus.fields.mark_fields_dirty` so Word
recalculates the cached results on first open.

::: docx_plus.bookmarks.crossref
    options:
      members:
        - add_cross_reference
        - CrossReferenceKind
        - NumberContext
