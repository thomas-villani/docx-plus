"""Anchored comments — body-side range markers plus comments.xml entries.

python-docx 1.x exposes a comments API but only writes the part-side
``<w:comment>``; it omits the three body-side OOXML elements that
actually anchor the comment to a text range. As a result, comments
added via python-docx show up in the review pane but have nothing in
the document text to attach to. This package writes the full set —
``w:commentRangeStart``, ``w:commentRangeEnd``, and the
``CommentReference`` marker run — alongside the comment body, so the
"show in document" link works in Word.

Comments are *threaded*: v0.4 added the ``commentsExtended.xml`` half of
the format, so a comment can carry replies and a resolved state the way
Word has modelled them since 2013.

Public surface:

- :func:`add_comment` — anchor a comment to a run, paragraph, or run range
- :func:`reply_to_comment` — attach a reply to an existing comment
- :func:`resolve_comment` / :func:`reopen_comment` — toggle a thread's
  resolved state
- :func:`read_comments` — list every comment with its anchored text
- :func:`read_threads` — the same comments grouped into threads
- :func:`edit_comment` — replace a comment's body text in place
- :func:`delete_comment` — remove a comment, its replies, and all anchors
- :func:`clear_all_comments` — scrub every comment from the document
- :class:`AnchoredComment` / :class:`CommentThread` — read-side result types
- :class:`CommentRef` — the write-side handle returned by ``add_comment``
- :class:`CommentIdRegistry` — pre-share across an editing session for
  many inserts

See SPEC §15 (the post-v0.1 roadmap) for where this capability was scoped.
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
from docx_plus.comments.threads import (
    CommentThread,
    read_threads,
    reopen_comment,
    reply_to_comment,
    resolve_comment,
)

__all__ = [
    "AnchoredComment",
    "CommentIdRegistry",
    "CommentNotFoundError",
    "CommentRef",
    "CommentTarget",
    "CommentThread",
    "add_comment",
    "clear_all_comments",
    "delete_comment",
    "edit_comment",
    "read_comments",
    "read_threads",
    "reopen_comment",
    "reply_to_comment",
    "resolve_comment",
]
