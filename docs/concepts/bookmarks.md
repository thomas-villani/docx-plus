# Bookmarks and cross-references

`bookmarks/anchor.py:add_bookmark(target, name, ...)` writes a paired
`<w:bookmarkStart w:id=N w:name=...>` / `<w:bookmarkEnd w:id=N/>`
around the target. Target shapes mirror `add_comment`: `Run`,
`Paragraph` (≥1 run), or `(Run, Run)` tuple. The name is validated
against Word's bookmark rules: `[A-Za-z_][A-Za-z0-9_]{0,39}`. Names
with spaces or punctuation are silently rejected by Word's UI but
accepted in raw OOXML, which leads to confusing failures —
`add_bookmark` raises eagerly instead.

For the calls, see the [bookmarks guide](../guides/bookmarks.md).

`delete_bookmark(doc, name)` removes every bookmark with the given
name (by name, not id, because that's what cross-references key off).
`read_bookmarks(doc)` returns a `BookmarkInfo` per bookmark with id,
name, anchored text, and paragraph index. `BookmarkIdRegistry` is the
fourth namespace (after SDT, comment, footnote / endnote each get
their own).

`bookmarks/crossref.py:add_cross_reference(paragraph, *, bookmark,
kind, hyperlink)` builds a `REF` (`kind="text"`) or `PAGEREF`
(`kind="page"`) [complex field](fields.md#complex-fields) via
`core.build_complex_field`. The `\h` flag is appended by default so Word
renders the cross-reference as a clickable link. Pair calls with
`mark_fields_dirty` so Word recalculates the cached results on first
open.

`core.build_bookmark` lives in `core/` rather than here so
[`publishing`](publishing.md) can make a caption referenceable — a `REF`
field can only point at a bookmark, never at the caption's `SEQ` field.
`BookmarkNameRegistry` (v0.5) guards duplicate names, which would make a
`REF` ambiguous, and mints the hidden `_Ref` + 9-digit anchors Word uses
for automatic cross-references.
