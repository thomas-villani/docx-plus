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

from lxml import etree

from docx_plus.core.oxml import el, sub
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
    registry = id_registry if id_registry is not None else FootnoteIdRegistry(
        _doc_for(paragraph)
    )
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
    registry = id_registry if id_registry is not None else EndnoteIdRegistry(
        _doc_for(paragraph)
    )
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

    p = sub(note, "w:p")
    p_pr = sub(p, "w:pPr")
    sub(p_pr, "w:pStyle", **{"w:val": para_style})

    # Leading reference glyph run.
    ref_run = sub(p, "w:r")
    ref_pr = sub(ref_run, "w:rPr")
    sub(ref_pr, "w:rStyle", **{"w:val": rstyle_ref})
    sub(ref_run, body_inner_ref_tag)

    # Body text run.
    text_run = sub(p, "w:r")
    text_t = sub(text_run, "w:t", **{"xml:space": "preserve"})
    text_t.text = text

    return note


__all__ = [
    "EndnoteRef",
    "FootnoteRef",
    "add_endnote",
    "add_footnote",
]
