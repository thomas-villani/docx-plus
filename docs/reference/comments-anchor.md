# `docx_plus.comments.anchor`

Anchor a comment to a run, paragraph, or run range — and undo the same.
Writes the three body-side OOXML elements python-docx skips
(`w:commentRangeStart`, `w:commentRangeEnd`, the `CommentReference`
marker run) plus the comment body in `comments.xml` (created on first
use). `delete_comment` is the inverse and idempotent.

Architecture walkthrough: [`ARCHITECTURE.md` §7.6](../ARCHITECTURE.md#76-anchored-comments).

::: docx_plus.comments.anchor
    options:
      members:
        - add_comment
        - edit_comment
        - delete_comment
        - clear_all_comments
        - CommentRef
        - CommentTarget
        - CommentNotFoundError
