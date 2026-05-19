"""Bookmark-id registry.

Bookmark ``w:id`` is its own uniqueness namespace, separate from SDT,
comment, and note ids. The body-side ``<w:bookmarkStart>`` /
``<w:bookmarkEnd>`` elements both carry the id on a direct ``@w:id``
attribute (not as a ``<w:id w:val=...>`` child like SDTs), so the
seeder uses :meth:`_collect_id_attrs`.

This module imports only from ``docx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx_plus.core.ids import _IdRegistryBase

if TYPE_CHECKING:
    from docx.document import Document


class BookmarkIdRegistry(_IdRegistryBase):
    """Tracks issued bookmark ``w:id`` values for one document-edit session."""

    def _seed_from_document(self, doc: Document) -> None:
        body = doc.element.body
        self._collect_id_attrs(body, ".//w:bookmarkStart")
        self._collect_id_attrs(body, ".//w:bookmarkEnd")


__all__ = ["BookmarkIdRegistry"]
