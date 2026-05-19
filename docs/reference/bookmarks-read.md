# `docx_plus.bookmarks.read`

Read every bookmark in a document. Each `BookmarkInfo` carries the
bookmark's id, name, the anchored text (what a `REF bookmark_name`
field would resolve to), and the paragraph index where the
`bookmarkStart` marker sits.

::: docx_plus.bookmarks.read
    options:
      members:
        - read_bookmarks
        - BookmarkInfo
