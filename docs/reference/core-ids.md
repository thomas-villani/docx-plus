# `docx_plus.core.ids`

Per-document id allocators. One registry per namespace per edit session —
OOXML reuses `w:id` across several disjoint uniqueness domains, and
bookmark id `7` does not collide with comment id `7`.

Most namespaces get their registry in the capability package that owns
them (`comments.CommentIdRegistry`, `notes.FootnoteIdRegistry`,
`numbering.NumIdRegistry`). Three live here instead:

- **`IdRegistry`** — SDT content-control ids, the original v0.1 case.
- **`ParaIdRegistry`** — `w14:paraId`, which is hex-rendered and unique
  across the whole *package* rather than within one part, because
  threaded comments key their parent/child links off it.
- **`BookmarkIdRegistry` / `BookmarkNameRegistry`** — moved here in v0.5
  because two capability packages need them: `bookmarks` owns
  `add_bookmark`, and `publishing` has to bookmark a caption to make it
  referenceable. SPEC §9.1 forbids the sibling import.
  Both are re-exported from
  [`docx_plus.bookmarks.registry`](bookmarks-registry.md).

Allocation comes in two flavours. `next()` mints a **random** 31-bit
value, which is right for an opaque handle. `next_sequential()` takes the
**lowest free** integer, which is what Word and python-docx do for
numbering — a `numbering.xml` full of nine-digit ids is needlessly
unreadable. `_MIN_ID` exists because `w:abstractNumId` legitimately
starts at 0, unlike every `w:id`.

::: docx_plus.core.ids
    options:
      members:
        - IdRegistry
        - ParaIdRegistry
        - BookmarkIdRegistry
        - BookmarkNameRegistry
        - DuplicateIdError
        - IdRangeError
        - DuplicateBookmarkNameError
