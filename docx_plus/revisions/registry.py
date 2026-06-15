"""Revision-id registry.

Unlike SDT, bookmark, comment, and note ids — each of which lives in its
own uniqueness namespace — *all* tracked-change revision elements share a
**single** ``w:id`` namespace. A ``<w:ins w:id="5">`` and a
``<w:del w:id="5">`` in the same document collide; so do a ``w:moveFrom``
and a ``w:rPrChange`` that reuse an id. This module ships a subclass of
:class:`~docx_plus.core.ids._IdRegistryBase` that seeds itself from every
revision-bearing element in the document body so :meth:`next` never
reissues an id already in use by any revision type.

The ``.//`` descendant axis deliberately reaches revision marks nested
inside ``w:pPr`` / ``w:rPr`` (paragraph-mark insertions and deletions),
not just the run-level ones, so those block reuse too.

This module imports only from ``docx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx_plus.core.ids import _IdRegistryBase

if TYPE_CHECKING:
    from docx.document import Document


# Every revision element that carries a ``w:id``. Move range markers
# (``*RangeStart`` / ``*RangeEnd``) carry the id too and must block reuse
# even when their wrapper or partner marker has been hand-edited away.
_REVISION_ID_TAGS: tuple[str, ...] = (
    ".//w:ins",
    ".//w:del",
    ".//w:moveFrom",
    ".//w:moveTo",
    ".//w:moveFromRangeStart",
    ".//w:moveFromRangeEnd",
    ".//w:moveToRangeStart",
    ".//w:moveToRangeEnd",
    ".//w:rPrChange",
    ".//w:pPrChange",
)


class RevisionIdRegistry(_IdRegistryBase):
    """Tracks issued revision ``w:id`` values for one document-edit session.

    All tracked-change element types draw from one shared id namespace, so
    this registry seeds from every revision tag in the body — run-level
    insertions/deletions, move wrappers and their range markers, and the
    property-change markers (including the paragraph-mark variants nested
    in ``w:pPr`` / ``w:rPr``).
    """

    def _seed_from_document(self, doc: Document) -> None:
        body = doc.element.body
        for expr in _REVISION_ID_TAGS:
            self._collect_id_attrs(body, expr)


__all__ = ["RevisionIdRegistry"]
