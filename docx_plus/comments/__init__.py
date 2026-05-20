"""Anchored comments — body-side range markers plus comments.xml entries.

python-docx 1.x exposes a comments API but only writes the part-side
``<w:comment>``; it omits the three body-side OOXML elements that
actually anchor the comment to a text range. As a result, comments
added via python-docx show up in the review pane but have nothing in
the document text to attach to. This package writes the full set —
``w:commentRangeStart``, ``w:commentRangeEnd``, and the
``CommentReference`` marker run — alongside the comment body, so the
"show in document" link works in Word.

Public surface:

- :func:`add_comment` — anchor a comment to a run, paragraph, or run range
- :func:`read_comments` — list every comment with its anchored text
- :func:`delete_comment` — remove a comment and all its anchors
- :class:`AnchoredComment` — the read-side result type
- :class:`CommentRef` — the write-side handle returned by ``add_comment``
- :class:`CommentIdRegistry` — pre-share across an editing session for
  many inserts

See SPEC §15 (deferred to v0.2) and ``notes-v0_1-scope.md §2.2`` for
context.
"""

from __future__ import annotations

from docx_plus.comments.anchor import (
    CommentNotFoundError,
    CommentRef,
    CommentTarget,
    add_comment,
    clear_all_comments,
    delete_comment,
    edit_comment,
)
from docx_plus.comments.read import AnchoredComment, read_comments
from docx_plus.comments.registry import CommentIdRegistry

__all__ = [
    "AnchoredComment",
    "CommentIdRegistry",
    "CommentNotFoundError",
    "CommentRef",
    "CommentTarget",
    "add_comment",
    "clear_all_comments",
    "delete_comment",
    "edit_comment",
    "read_comments",
]
