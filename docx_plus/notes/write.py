"""Insert footnotes and endnotes.

Footnotes and endnotes share the same shape: a reference marker
``<w:r><w:rPr><w:rStyle val=".."/></w:rPr><w:footnoteReference|w:endnoteReference w:id=N/></w:r>``
in the body, plus a content entry in the corresponding separate part
(``word/footnotes.xml`` or ``word/endnotes.xml``).

Insert-only is sufficient for v0.2 — in-place edits of existing notes
are deferred to v0.3.

This module imports only from ``docx_plus.core`` and the sibling
``docx_plus.notes.registry`` (SPEC §9.1).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, cast

from docx.opc.part import XmlPart
from lxml import etree

from docx_plus.core import DocxPlusError
from docx_plus.core.oxml import el, remove, sub, xpath
from docx_plus.core.parts import (
    ENDNOTES_SPEC,
    FOOTNOTES_SPEC,
    PartSpec,
    get_or_create_part,
)
from docx_plus.notes.registry import (
    EndnoteIdRegistry,
    FootnoteIdRegistry,
    _NoteIdRegistryBase,
)

if TYPE_CHECKING:
    from docx.document import Document
    from docx.text.paragraph import Paragraph


class NoteNotFoundError(DocxPlusError, KeyError):
    """Raised when no footnote / endnote with the requested id exists.

    Subclasses :class:`KeyError` so existing ``except KeyError:`` clauses
    still catch it; also :class:`DocxPlusError` per SPEC §9.7.
    """


@dataclass(frozen=True)
class FootnoteRef:
    """Handle for an inserted footnote."""

    note_id: int
    body_element: etree._Element


@dataclass(frozen=True)
class EndnoteRef:
    """Handle for an inserted endnote."""

    note_id: int
    body_element: etree._Element


def add_footnote(
    paragraph: Paragraph,
    text: str,
    *,
    id_registry: FootnoteIdRegistry | None = None,
) -> FootnoteRef:
    """Append a footnote reference to ``paragraph`` and a footnote body to ``footnotes.xml``.

    Args:
        paragraph: A python-docx :class:`~docx.text.paragraph.Paragraph`
            in the main document body. The reference marker run is
            appended after the paragraph's existing runs.
        text: Footnote body text. Whitespace is preserved.
        id_registry: Pre-existing :class:`FootnoteIdRegistry` to share
            across an editing session.

    Returns:
        A :class:`FootnoteRef` with the assigned note id and a handle
        to the note body element.

    Example:
        >>> from docx import Document
        >>> from docx_plus.notes import add_footnote
        >>> doc = Document()
        >>> p = doc.add_paragraph("See the footnote.")
        >>> add_footnote(p, "Explanatory text.")
    """
    registry = id_registry if id_registry is not None else FootnoteIdRegistry(_doc_for(paragraph))
    note_id, note_body = _add_note(
        paragraph,
        text,
        spec=FOOTNOTES_SPEC,
        ref_tag="w:footnoteReference",
        body_tag="w:footnote",
        rstyle_ref="FootnoteReference",
        body_inner_ref_tag="w:footnoteRef",
        para_style="FootnoteText",
        registry=registry,
    )
    return FootnoteRef(note_id=note_id, body_element=note_body)


def add_endnote(
    paragraph: Paragraph,
    text: str,
    *,
    id_registry: EndnoteIdRegistry | None = None,
) -> EndnoteRef:
    """Append an endnote reference to ``paragraph`` and an endnote body to ``endnotes.xml``.

    Same shape as :func:`add_footnote` but writes the endnote variants
    of the reference elements and the endnotes part.

    Args:
        paragraph: A python-docx :class:`~docx.text.paragraph.Paragraph`
            in the main document body.
        text: Endnote body text.
        id_registry: Pre-existing :class:`EndnoteIdRegistry`.

    Returns:
        A :class:`EndnoteRef` with the assigned note id and a handle
        to the note body element.
    """
    registry = id_registry if id_registry is not None else EndnoteIdRegistry(_doc_for(paragraph))
    note_id, note_body = _add_note(
        paragraph,
        text,
        spec=ENDNOTES_SPEC,
        ref_tag="w:endnoteReference",
        body_tag="w:endnote",
        rstyle_ref="EndnoteReference",
        body_inner_ref_tag="w:endnoteRef",
        para_style="EndnoteText",
        registry=registry,
    )
    return EndnoteRef(note_id=note_id, body_element=note_body)


def edit_footnote(doc: Document, note_id: int, text: str) -> None:
    """Replace the body text of an existing footnote in place.

    Removes all child paragraphs of the matching ``<w:footnote>`` element
    and appends a fresh paragraph with ``text``. The reference glyph run
    is rebuilt as part of the new paragraph; the body-side reference
    marker run in the main document is untouched.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to mutate.
        note_id: The ``w:id`` of the footnote to edit. Must be ``>= 1``;
            ids ``-1`` and ``0`` are reserved separator entries and are
            not editable.
        text: New footnote body text. Whitespace is preserved.

    Raises:
        ValueError: If ``note_id`` is ``<= 0``.
        NoteNotFoundError: If no footnote with ``note_id`` exists,
            including the case where the footnotes part is absent.
    """
    _edit_note(
        doc,
        note_id,
        text,
        spec=FOOTNOTES_SPEC,
        body_tag="w:footnote",
        rstyle_ref="FootnoteReference",
        body_inner_ref_tag="w:footnoteRef",
        para_style="FootnoteText",
    )


def edit_endnote(doc: Document, note_id: int, text: str) -> None:
    """Replace the body text of an existing endnote in place.

    Same shape as :func:`edit_footnote` but targets the endnotes part.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to mutate.
        note_id: The ``w:id`` of the endnote. Must be ``>= 1``.
        text: New endnote body text.

    Raises:
        ValueError: If ``note_id`` is ``<= 0``.
        NoteNotFoundError: If no endnote with ``note_id`` exists.
    """
    _edit_note(
        doc,
        note_id,
        text,
        spec=ENDNOTES_SPEC,
        body_tag="w:endnote",
        rstyle_ref="EndnoteReference",
        body_inner_ref_tag="w:endnoteRef",
        para_style="EndnoteText",
    )


def _edit_note(  # noqa: PLR0913
    doc: Document,
    note_id: int,
    text: str,
    *,
    spec: PartSpec,
    body_tag: str,
    rstyle_ref: str,
    body_inner_ref_tag: str,
    para_style: str,
) -> None:
    """Shared in-place edit for footnote / endnote bodies."""
    if note_id <= 0:
        raise ValueError(
            f"note ids -1 and 0 are reserved separator entries and are not editable; got {note_id}"
        )

    try:
        part = cast("XmlPart", doc.part.part_related_by(spec.relationship_type))
    except KeyError as exc:
        raise NoteNotFoundError(note_id) from exc

    matches = xpath(part.element, f"./{body_tag}[@w:id=$nid]", nid=str(note_id))
    if not matches:
        raise NoteNotFoundError(note_id)

    note_el = matches[0]
    # Strip ALL block-level children — note bodies can legally contain
    # tables and SDTs in addition to paragraphs (same reason as
    # edit_comment, see ECMA-376 17.13.4.2 / footnote equivalents).
    # The note element's `w:id` / `w:type` attributes live on the
    # element itself, so removal preserves them.
    for child in list(note_el):
        remove(child)
    note_el.append(
        _build_note_paragraph(
            text,
            rstyle_ref=rstyle_ref,
            body_inner_ref_tag=body_inner_ref_tag,
            para_style=para_style,
        )
    )


def _add_note(  # noqa: PLR0913
    paragraph: Paragraph,
    text: str,
    *,
    spec: PartSpec,
    ref_tag: str,
    body_tag: str,
    rstyle_ref: str,
    body_inner_ref_tag: str,
    para_style: str,
    registry: _NoteIdRegistryBase,
) -> tuple[int, etree._Element]:
    """Shared core for footnote / endnote insertion."""
    note_id = registry.next()
    nid = str(note_id)

    # Body-side reference run.
    ref_run = el("w:r")
    ref_pr = sub(ref_run, "w:rPr")
    sub(ref_pr, "w:rStyle", **{"w:val": rstyle_ref})
    sub(ref_run, ref_tag, **{"w:id": nid})
    paragraph._p.append(ref_run)

    # Part-side note body.
    document = _doc_for(paragraph)
    _, root = get_or_create_part(document, spec)
    note_body = _build_note_body(
        note_id=note_id,
        text=text,
        body_tag=body_tag,
        rstyle_ref=rstyle_ref,
        body_inner_ref_tag=body_inner_ref_tag,
        para_style=para_style,
    )
    root.append(note_body)

    return note_id, note_body


def _doc_for(paragraph: Paragraph) -> Document:
    """Return the :class:`Document` containing ``paragraph``."""
    part = paragraph.part
    document = getattr(part, "document", None)
    if document is None:
        raise ValueError(
            "add_footnote / add_endnote only support the main document body "
            f"in v0.2; got paragraph parented to {type(part).__name__}"
        )
    return cast("Document", document)


def _build_note_body(
    *,
    note_id: int,
    text: str,
    body_tag: str,
    rstyle_ref: str,
    body_inner_ref_tag: str,
    para_style: str,
) -> etree._Element:
    note = el(body_tag, **{"w:id": str(note_id)})
    note.append(
        _build_note_paragraph(
            text,
            rstyle_ref=rstyle_ref,
            body_inner_ref_tag=body_inner_ref_tag,
            para_style=para_style,
        )
    )
    return note


def _build_note_paragraph(
    text: str,
    *,
    rstyle_ref: str,
    body_inner_ref_tag: str,
    para_style: str,
) -> etree._Element:
    """Build a note-body ``<w:p>`` with the reference glyph and text run.

    Shared by :func:`add_footnote` / :func:`add_endnote` (initial
    insertion) and :func:`edit_footnote` / :func:`edit_endnote`
    (in-place replacement).
    """
    p = el("w:p")
    p_pr = sub(p, "w:pPr")
    sub(p_pr, "w:pStyle", **{"w:val": para_style})

    ref_run = sub(p, "w:r")
    ref_pr = sub(ref_run, "w:rPr")
    sub(ref_pr, "w:rStyle", **{"w:val": rstyle_ref})
    sub(ref_run, body_inner_ref_tag)

    text_run = sub(p, "w:r")
    text_t = sub(text_run, "w:t", **{"xml:space": "preserve"})
    text_t.text = text

    return p


__all__ = [
    "EndnoteRef",
    "FootnoteRef",
    "NoteNotFoundError",
    "add_endnote",
    "add_footnote",
    "edit_endnote",
    "edit_footnote",
]
