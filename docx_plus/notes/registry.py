"""Footnote / endnote id registries.

Footnote and endnote ids are two separate uniqueness namespaces. Ids
``-1`` and ``0`` are reserved by Word for the *separator* and
*continuation separator* entries respectively, so both registries
refuse to issue those.

This module imports only from ``docx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import XmlPart

from docx_plus.core.ids import _IdRegistryBase

if TYPE_CHECKING:
    from docx.document import Document


class _NoteIdRegistryBase(_IdRegistryBase):
    """Common seeding logic for footnote and endnote registries.

    Subclasses set ``_RELATIONSHIP_TYPE`` and ``_NOTE_TAG`` (the part
    root's child name, e.g. ``"./w:footnote"``).
    """

    _RELATIONSHIP_TYPE: str = ""
    _NOTE_TAG: str = ""

    def _seed_from_document(self, doc: Document) -> None:
        # Ids -1 and 0 are reserved by Word (separator, continuationSeparator).
        # ``next()`` already clamps to >= 1 and ``reserve()`` already rejects
        # values outside [1, 2**31 - 1], so no pre-seeding is needed to block
        # them — the range check fires first.
        document_part = doc.part
        try:
            part = cast("XmlPart", document_part.part_related_by(self._RELATIONSHIP_TYPE))
        except KeyError:
            return
        self._collect_id_attrs(part.element, self._NOTE_TAG)


class FootnoteIdRegistry(_NoteIdRegistryBase):
    """Tracks issued footnote ``w:id`` values for one document-edit session."""

    _RELATIONSHIP_TYPE = RT.FOOTNOTES
    _NOTE_TAG = "./w:footnote"


class EndnoteIdRegistry(_NoteIdRegistryBase):
    """Tracks issued endnote ``w:id`` values for one document-edit session."""

    _RELATIONSHIP_TYPE = RT.ENDNOTES
    _NOTE_TAG = "./w:endnote"


__all__ = ["EndnoteIdRegistry", "FootnoteIdRegistry"]
