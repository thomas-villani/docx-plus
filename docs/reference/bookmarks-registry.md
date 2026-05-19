# `docx_plus.bookmarks.registry`

Per-document registry of issued bookmark `w:id` values. Bookmark ids
live in their own uniqueness namespace, separate from SDT, comment, and
note ids. Body-side `<w:bookmarkStart>` / `<w:bookmarkEnd>` elements
both carry the id on a direct `@w:id` attribute, so the seeder uses
the attribute-form collector inherited from `_IdRegistryBase`.

::: docx_plus.bookmarks.registry
    options:
      members:
        - BookmarkIdRegistry
