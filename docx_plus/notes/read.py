"""Read footnotes / endnotes from a document.

Each note is paired with the paragraph index of its reference marker
in the body so callers can locate where the note is referenced.

This module imports only from ``docx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, cast

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import XmlPart
from lxml import etree

from docx_plus.core.ns import qn
from docx_plus.core.oxml import xpath

if TYPE_CHECKING:
    from docx.document import Document


@dataclass(frozen=True)
class NoteContent:
    """A footnote or endnote with its body text and reference location.

    Attributes:
        note_id: The ``w:id`` value of the note.
        text: The note body text. Reserved separator entries (ids
            ``-1`` and ``0``) are filtered out before this list is
            built, so callers see only user-authored notes.
        paragraph_index: Zero-based index (within ``doc.paragraphs``)
            of the paragraph holding the reference marker. ``-1`` if
            the note is in the part but no body reference exists.
    """

    note_id: int
    text: str
    paragraph_index: int


def read_footnotes(doc: Document) -> list[NoteContent]:
    """Return every user-authored footnote in ``doc``.

    Reserved entries (separator and continuation separator, ids ``-1``
    and ``0``) are filtered out.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to scan.

    Returns:
        One :class:`NoteContent` per footnote, in part order. ``[]`` if
        the document has no footnotes part.
    """
    return _read_notes(
        doc,
        relationship_type=RT.FOOTNOTES,
        note_tag="./w:footnote",
        ref_tag=".//w:footnoteReference",
    )


def read_endnotes(doc: Document) -> list[NoteContent]:
    """Return every user-authored endnote in ``doc``.

    Same contract as :func:`read_footnotes`.
    """
    return _read_notes(
        doc,
        relationship_type=RT.ENDNOTES,
        note_tag="./w:endnote",
        ref_tag=".//w:endnoteReference",
    )


def _read_notes(
    doc: Document,
    *,
    relationship_type: str,
    note_tag: str,
    ref_tag: str,
) -> list[NoteContent]:
    try:
        part = cast("XmlPart", doc.part.part_related_by(relationship_type))
    except KeyError:
        return []

    root = part.element
    body = doc.element.body
    paragraph_elements = list(xpath(body, ".//w:p"))

    # Map note id → paragraph index of the body-side reference marker.
    ref_paragraph: dict[str, int] = {}
    for ref in xpath(body, ref_tag):
        nid = ref.get(qn("w:id"))
        if nid is None:
            continue
        ancestor = ref.getparent()
        while ancestor is not None and ancestor.tag != qn("w:p"):
            ancestor = ancestor.getparent()
        if ancestor is None:
            continue
        try:
            ref_paragraph[nid] = paragraph_elements.index(ancestor)
        except ValueError:
            continue

    result: list[NoteContent] = []
    for note_el in xpath(root, note_tag):
        nid_raw = note_el.get(qn("w:id"))
        if nid_raw is None:
            continue
        try:
            nid = int(nid_raw)
        except ValueError:
            continue
        if nid <= 0:
            # Reserved: -1 (separator), 0 (continuation separator).
            continue

        # Two-tier filter (L16): reserved entries normally use ids <= 0, but a
        # tool may also tag a separator with a positive id. The @w:type check
        # is deliberately belt-and-suspenders — note that it would also drop a
        # (highly unusual) user note that set w:type to one of the two
        # separator values. Any other legal w:type passes through.
        note_type = note_el.get(qn("w:type"))
        if note_type in {"separator", "continuationSeparator"}:
            continue

        text = _note_body_text(note_el)
        para_idx = ref_paragraph.get(nid_raw, -1)
        result.append(
            NoteContent(note_id=nid, text=text, paragraph_index=para_idx)
        )
    return result


def _note_body_text(note_el: etree._Element) -> str:
    """Concatenate every ``<w:t>`` text inside the note body."""
    parts: list[str] = []
    for t in xpath(note_el, ".//w:t"):
        if t.text:
            parts.append(t.text)
    return "".join(parts)


__all__ = ["NoteContent", "read_endnotes", "read_footnotes"]
