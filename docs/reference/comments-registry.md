# `docx_plus.comments.registry`

Per-document registry of issued comment `w:id` values. Comment ids live
in a separate uniqueness namespace from SDT, bookmark, and note ids —
comment `5` does not collide with bookmark `5`. The registry seeds
itself from both the comments part and any orphaned body-side range
markers so partially-deleted comments can't trigger id reuse.

::: docx_plus.comments.registry
    options:
      members:
        - CommentIdRegistry
