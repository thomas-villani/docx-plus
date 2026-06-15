"""Enumerate every tracked change in a document.

Inverse of :func:`docx_plus.revisions.mark_insertion` /
:func:`~docx_plus.revisions.mark_deletion`, and the reader for revision
marks Word itself authored: walks the document body once and reports every
``w:ins``, ``w:del``, ``w:moveFrom`` / ``w:moveTo``, ``w:rPrChange``, and
``w:pPrChange`` with its id, author, timestamp, type, and affected text.

Run text inside revision wrappers is invisible to python-docx's
``paragraph.runs``, so all text is read through our own XPath — insertions
from ``<w:t>`` and deletions from ``<w:delText>``.

This module imports only from ``docx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass
from typing import TYPE_CHECKING, Literal

from lxml import etree

from docx_plus.core.ns import qn
from docx_plus.core.oxml import xpath

if TYPE_CHECKING:
    from docx.document import Document


RevisionType = Literal[
    "insertion",
    "deletion",
    "move_from",
    "move_to",
    "format_run",
    "format_paragraph",
    "paragraph_mark_insertion",
    "paragraph_mark_deletion",
]


@dataclass(frozen=True)
class TrackedChange:
    """One revision mark paired with the text it affects.

    Attributes:
        revision_id: The ``w:id`` value of the revision element.
        revision_type: One of the :data:`RevisionType` literals.
        author: The ``w:author`` attribute (may be empty).
        timestamp: The ``w:date`` attribute parsed as a timezone-aware UTC
            :class:`datetime`, or ``None`` if absent or unparseable.
        text: For insertions, the inserted ``<w:t>`` text. For deletions,
            the deleted ``<w:delText>`` text. For moves, the moved run text.
            Empty for format changes and paragraph-mark revisions (the mark
            itself carries no text).
        paragraph_index: Zero-based index (within ``doc.paragraphs``) of the
            paragraph containing the revision element, or ``-1`` if it could
            not be resolved.
    """

    revision_id: int
    revision_type: RevisionType
    author: str
    timestamp: dt.datetime | None
    text: str
    paragraph_index: int


def read_revisions(doc: Document) -> list[TrackedChange]:
    """Return every tracked change in ``doc`` in document order.

    Enumerates run-level insertions/deletions, move source/destination
    wrappers, run- and paragraph-property changes, and paragraph-mark
    insertions/deletions. Move *range markers* (the bookmark-like
    ``*RangeStart`` / ``*RangeEnd`` delimiters) are not reported as separate
    entries — the ``w:moveFrom`` / ``w:moveTo`` wrapper that carries the
    moved text and metadata is.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to scan.

    Returns:
        One :class:`TrackedChange` per revision element, in document order.
        Returns ``[]`` for a document with no tracked changes.
    """
    body = doc.element.body
    paragraph_elements = list(xpath(body, ".//w:p"))

    handlers = {
        qn("w:ins"): _read_ins,
        qn("w:del"): _read_del,
        qn("w:moveFrom"): _read_move_from,
        qn("w:moveTo"): _read_move_to,
        qn("w:rPrChange"): _read_rpr_change,
        qn("w:pPrChange"): _read_ppr_change,
    }

    result: list[TrackedChange] = []
    for elem in body.iter():
        handler = handlers.get(elem.tag)
        if handler is None:
            continue
        change = handler(elem, paragraph_elements)
        if change is not None:
            result.append(change)
    return result


# ---------------------------------------------------------------------------
# Per-element readers.
# ---------------------------------------------------------------------------


def _read_ins(
    elem: etree._Element, paragraphs: list[etree._Element]
) -> TrackedChange | None:
    rtype: RevisionType = "paragraph_mark_insertion" if _is_property_mark(elem) else "insertion"
    text = "" if rtype == "paragraph_mark_insertion" else _collect_text(elem, (".//w:t",))
    return _build(elem, rtype, text, paragraphs)


def _read_del(
    elem: etree._Element, paragraphs: list[etree._Element]
) -> TrackedChange | None:
    rtype: RevisionType = "paragraph_mark_deletion" if _is_property_mark(elem) else "deletion"
    text = "" if rtype == "paragraph_mark_deletion" else _collect_text(elem, (".//w:delText",))
    return _build(elem, rtype, text, paragraphs)


def _read_move_from(
    elem: etree._Element, paragraphs: list[etree._Element]
) -> TrackedChange | None:
    return _build(elem, "move_from", _collect_text(elem, (".//w:t", ".//w:delText")), paragraphs)


def _read_move_to(
    elem: etree._Element, paragraphs: list[etree._Element]
) -> TrackedChange | None:
    return _build(elem, "move_to", _collect_text(elem, (".//w:t", ".//w:delText")), paragraphs)


def _read_rpr_change(
    elem: etree._Element, paragraphs: list[etree._Element]
) -> TrackedChange | None:
    return _build(elem, "format_run", "", paragraphs)


def _read_ppr_change(
    elem: etree._Element, paragraphs: list[etree._Element]
) -> TrackedChange | None:
    return _build(elem, "format_paragraph", "", paragraphs)


# ---------------------------------------------------------------------------
# Shared helpers.
# ---------------------------------------------------------------------------


def _build(
    elem: etree._Element,
    rtype: RevisionType,
    text: str,
    paragraphs: list[etree._Element],
) -> TrackedChange | None:
    """Assemble a :class:`TrackedChange`, or ``None`` if ``w:id`` is unparseable."""
    raw = elem.get(qn("w:id"))
    if raw is None:
        return None
    try:
        revision_id = int(raw)
    except ValueError:
        return None

    return TrackedChange(
        revision_id=revision_id,
        revision_type=rtype,
        author=elem.get(qn("w:author")) or "",
        timestamp=_parse_date(elem.get(qn("w:date"))),
        text=text,
        paragraph_index=_paragraph_index(elem, paragraphs),
    )


def _is_property_mark(elem: etree._Element) -> bool:
    """True if ``elem`` is a paragraph/run-mark revision, not run-level content.

    A run-level ``w:ins`` / ``w:del`` parents runs directly under a ``w:p``
    (or another revision). A paragraph-mark revision sits *inside* the
    paragraph or run properties — ``<w:pPr><w:rPr><w:ins/></w:rPr></w:pPr>``
    — so its parent is ``w:rPr`` or ``w:pPr``.
    """
    parent = elem.getparent()
    if parent is None:
        return False
    return parent.tag in (qn("w:rPr"), qn("w:pPr"))


def _collect_text(elem: etree._Element, exprs: tuple[str, ...]) -> str:
    """Concatenate text from every node matched by ``exprs`` in document order."""
    parts: list[str] = []
    for expr in exprs:
        for node in xpath(elem, expr):
            if node.text:
                parts.append(node.text)
    return "".join(parts)


def _paragraph_index(elem: etree._Element, paragraphs: list[etree._Element]) -> int:
    """Index into ``paragraphs`` of the ``w:p`` ancestor of ``elem``, or ``-1``."""
    ancestor = elem.getparent()
    p_tag = qn("w:p")
    while ancestor is not None and ancestor.tag != p_tag:
        ancestor = ancestor.getparent()
    if ancestor is None:
        return -1
    try:
        return paragraphs.index(ancestor)
    except ValueError:
        return -1


def _parse_date(value: str | None) -> dt.datetime | None:
    """Parse an ``xsd:dateTime`` value, tolerating the trailing ``Z`` form."""
    if not value:
        return None
    normalized = value.rstrip("Z")
    if value.endswith("Z"):
        normalized = normalized + "+00:00"
    try:
        return dt.datetime.fromisoformat(normalized)
    except ValueError:
        return None


__all__ = ["RevisionType", "TrackedChange", "read_revisions"]
