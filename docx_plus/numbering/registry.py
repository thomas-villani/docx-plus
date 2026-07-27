"""Numbering id registries — ``w:numId`` and ``w:abstractNumId``.

``numbering.xml`` carries two disjoint id namespaces. A ``w:abstractNum``
is the *definition* — what the levels look like — and a ``w:num`` is an
*instance* of one, which is what a paragraph's ``w:numPr`` actually
references. Two paragraphs sharing a ``numId`` continue one sequence;
two ``num`` entries pointing at the same ``abstractNumId`` are
independent sequences with identical formatting, which is how Word
restarts a list.

Both registries allocate with :meth:`~docx_plus.core.ids._IdRegistryBase.next_sequential`
rather than the random :meth:`~docx_plus.core.ids._IdRegistryBase.next`
every other namespace uses. Word and python-docx both number lists with
the lowest free integer, and a ``numbering.xml`` full of nine-digit ids
is needlessly unreadable — these ids are read by humans debugging list
behaviour far more often than most.

This module imports only from ``docx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import XmlPart

from docx_plus.core.ids import _IdRegistryBase

if TYPE_CHECKING:
    from docx.document import Document


class _NumberingIdRegistryBase(_IdRegistryBase):
    """Common seeding logic for the two ``numbering.xml`` namespaces.

    Subclasses set ``_ENTRY_TAG`` (the root's child name) and
    ``_ID_ATTR`` (the attribute carrying the id).
    """

    _ENTRY_TAG: str = ""
    _ID_ATTR: str = ""

    def _seed_from_document(self, doc: Document) -> None:
        # Read-only: never fabricates the part. A document with no
        # numbering.xml simply has an empty namespace, and constructing a
        # registry must not have the side effect of creating a part.
        try:
            part = cast("XmlPart", doc.part.part_related_by(RT.NUMBERING))
        except KeyError:
            return
        self._collect_named_attrs(part.element, self._ENTRY_TAG, self._ID_ATTR)


class NumIdRegistry(_NumberingIdRegistryBase):
    """Tracks issued ``w:num`` ids for one document-edit session.

    ``w:numId`` is what a paragraph's ``w:numPr`` references. Note that
    ``numId`` ``0`` is not an entry id — inside a ``w:numPr`` it is the
    sentinel meaning "no numbering", which is how a paragraph opts out of
    a list its *style* applies. Allocation therefore starts at 1.
    """

    _ENTRY_TAG = "./w:num"
    _ID_ATTR = "w:numId"


class AbstractNumIdRegistry(_NumberingIdRegistryBase):
    """Tracks issued ``w:abstractNum`` ids for one document-edit session.

    Unlike every other id namespace in the library, ``0`` is legal here —
    python-docx's own bundled template ships ``abstractNumId`` 0 through
    8 — so this registry lowers :attr:`_MIN_ID` accordingly.
    """

    _ENTRY_TAG = "./w:abstractNum"
    _ID_ATTR = "w:abstractNumId"
    _MIN_ID = 0


__all__ = ["AbstractNumIdRegistry", "NumIdRegistry"]
