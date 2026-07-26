"""Bookmark id and name registries.

Bookmark ``w:id`` is its own uniqueness namespace, separate from SDT,
comment, and note ids. Bookmarks are also the one thing in the format
addressed by *name*, which is a second namespace needing its own
allocator.

Both classes moved to :mod:`docx_plus.core.ids` in v0.5 and are
re-exported here, so ``from docx_plus.bookmarks import BookmarkIdRegistry``
is unchanged. The move was forced by SPEC §9.1: ``publishing`` has to
bookmark a caption to make it referenceable — a ``REF`` field can only
point at a bookmark, never at the caption's own ``SEQ`` field — and it
cannot import from a sibling capability to do it.

This module imports only from ``docx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from docx_plus.core.ids import (
    BookmarkIdRegistry,
    BookmarkNameRegistry,
    DuplicateBookmarkNameError,
)

__all__ = [
    "BookmarkIdRegistry",
    "BookmarkNameRegistry",
    "DuplicateBookmarkNameError",
]
