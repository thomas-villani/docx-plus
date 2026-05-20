"""Bookmarks and cross-references — paired body markers plus REF/PAGEREF.

OOXML bookmarks are tiny: a ``<w:bookmarkStart>`` and ``<w:bookmarkEnd>``
pair with matching ``w:id`` and a ``w:name``. Cross-references are
complex fields keyed off the bookmark name (``REF`` for text,
``PAGEREF`` for page number). python-docx exposes neither.

Public surface:

- :func:`add_bookmark` — anchor a bookmark to a run, paragraph, or run
  range
- :func:`delete_bookmark` — remove by name
- :func:`read_bookmarks` — list every bookmark with its anchored text
- :func:`add_cross_reference` — insert a ``REF`` or ``PAGEREF`` field
- :class:`BookmarkInfo` — the read-side result
- :class:`BookmarkRef` — the write-side handle
- :class:`BookmarkIdRegistry` — pre-share across an editing session

See SPEC §15 (the post-v0.1 roadmap) for where this capability was scoped.
"""

from __future__ import annotations

from docx_plus.bookmarks.anchor import (
    BookmarkRef,
    BookmarkTarget,
    add_bookmark,
    delete_bookmark,
)
from docx_plus.bookmarks.crossref import CrossReferenceKind, add_cross_reference
from docx_plus.bookmarks.read import BookmarkInfo, read_bookmarks
from docx_plus.bookmarks.registry import BookmarkIdRegistry

__all__ = [
    "BookmarkIdRegistry",
    "BookmarkInfo",
    "BookmarkRef",
    "BookmarkTarget",
    "CrossReferenceKind",
    "add_bookmark",
    "add_cross_reference",
    "delete_bookmark",
    "read_bookmarks",
]
