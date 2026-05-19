# `docx_plus.bookmarks.anchor`

Paired `<w:bookmarkStart>` / `<w:bookmarkEnd>` body markers with a
shared `w:id` and a human-readable `w:name`. Cross-references key off
the name. python-docx provides no abstraction for bookmarks; this
module fills the gap, validating names against Word's rules
(`[A-Za-z_][A-Za-z0-9_]{0,39}`) so silently broken cross-references
become impossible.

Architecture walkthrough: [`ARCHITECTURE.md` §7.8](../ARCHITECTURE.md#78-bookmarks-and-cross-references).

::: docx_plus.bookmarks.anchor
    options:
      members:
        - add_bookmark
        - delete_bookmark
        - BookmarkRef
        - BookmarkTarget
