# `docx_plus.bookmarks.registry`

Per-document registries for the two bookmark namespaces.

`BookmarkIdRegistry` tracks issued `w:id` values, which live in their own
uniqueness namespace separate from SDT, comment, and note ids. Body-side
`<w:bookmarkStart>` / `<w:bookmarkEnd>` elements both carry the id on a
direct `@w:id` attribute, so the seeder uses the attribute-form collector
inherited from `_IdRegistryBase`.

`BookmarkNameRegistry` (v0.5) tracks bookmark *names*. Bookmarks are the
one thing in the format addressed by name, and nothing stops a document
carrying two with the same `w:name` — a duplicate makes a `REF`
ambiguous and makes `delete_bookmark` remove both. It also mints hidden
Word-style anchors with `next_ref_name()`, in the `_Ref` + 9-digit form
Word itself uses for auto-generated cross-reference targets. The leading
underscore is load-bearing: Word omits underscore-prefixed bookmarks from
its Bookmark dialog, so machine-generated caption anchors stay out of the
user's list.

Both classes live in
[`docx_plus.core.ids`](core-ids.md) as of v0.5 and are re-exported here.
The move was forced by SPEC §9.1 — `publishing` has to bookmark a caption
to make it referenceable and cannot import from a sibling capability to
do it.

::: docx_plus.bookmarks.registry
    options:
      members:
        - BookmarkIdRegistry
        - BookmarkNameRegistry
        - DuplicateBookmarkNameError
