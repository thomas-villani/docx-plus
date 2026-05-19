"""Read every anchored comment from a document.

Inverse of :func:`docx_plus.comments.add_comment`: walks the comments
part and pairs each ``<w:comment>`` with the body-side range it anchors,
extracting the comment text *and* the document text the comment is
attached to.

This module imports only from ``docx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

import datetime as dt
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
class AnchoredComment:
    """A comment paired with the document text it anchors to.

    Attributes:
        comment_id: The ``w:id`` value of the comment.
        author: The ``w:author`` attribute (may be empty).
        initials: The ``w:initials`` attribute, or ``None`` if absent.
        timestamp: The ``w:date`` attribute parsed as a timezone-aware
            UTC :class:`datetime`, or ``None`` if the attribute is
            absent or unparseable.
        text: The comment body text. Multiple text runs are concatenated.
        anchored_text: The document text between the comment's
            ``commentRangeStart`` and ``commentRangeEnd`` markers.
            Empty if no matching body range exists (orphaned comment).
        paragraph_index: Zero-based index (within
            ``doc.paragraphs``) of the paragraph that contains the
            ``commentRangeStart`` marker. ``-1`` for orphaned comments.
    """

    comment_id: int
    author: str
    initials: str | None
    timestamp: dt.datetime | None
    text: str
    anchored_text: str
    paragraph_index: int


def read_comments(doc: Document) -> list[AnchoredComment]:
    """Return every comment in ``doc`` paired with the text it anchors to.

    A comment with no matching body range still appears in the result
    with ``anchored_text=""`` and ``paragraph_index=-1`` — this is the
    "orphaned" state that python-docx's ``add_comment`` produces.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to scan.

    Returns:
        One :class:`AnchoredComment` per comment, in
        ``comments.xml`` order. Returns ``[]`` if the document has no
        comments part at all.
    """
    try:
        comments_part = cast("XmlPart", doc.part.part_related_by(RT.COMMENTS))
    except KeyError:
        return []

    comments_root = comments_part.element
    body = doc.element.body
    paragraph_elements = list(xpath(body, ".//w:p"))

    result: list[AnchoredComment] = []
    for comment_el in xpath(comments_root, "./w:comment"):
        cid_raw = comment_el.get(qn("w:id"))
        if cid_raw is None:
            continue
        try:
            cid = int(cid_raw)
        except ValueError:
            continue

        author = comment_el.get(qn("w:author")) or ""
        initials = comment_el.get(qn("w:initials"))
        timestamp = _parse_date(comment_el.get(qn("w:date")))

        text = _comment_body_text(comment_el)

        anchored_text, paragraph_index = _anchor_lookup(
            body, paragraph_elements, str(cid)
        )

        result.append(
            AnchoredComment(
                comment_id=cid,
                author=author,
                initials=initials,
                timestamp=timestamp,
                text=text,
                anchored_text=anchored_text,
                paragraph_index=paragraph_index,
            )
        )
    return result


def _comment_body_text(comment_el: etree._Element) -> str:
    """Concat every ``<w:t>`` inside the comment body in document order."""
    parts: list[str] = []
    for t in xpath(comment_el, ".//w:t"):
        if t.text:
            parts.append(t.text)
    return "".join(parts)


def _anchor_lookup(
    body: etree._Element,
    paragraph_elements: list[etree._Element],
    cid: str,
) -> tuple[str, int]:
    """Return ``(anchored_text, paragraph_index)`` for comment ``cid``."""
    starts = xpath(body, ".//w:commentRangeStart[@w:id=$cid]", cid=cid)
    ends = xpath(body, ".//w:commentRangeEnd[@w:id=$cid]", cid=cid)
    if not starts or not ends:
        return ("", -1)

    start = starts[0]
    end = ends[0]
    text = _text_between(body, start, end)

    # paragraph_index is the index in body's w:p list that contains start.
    paragraph_index = -1
    ancestor = start.getparent()
    while ancestor is not None and ancestor.tag != qn("w:p"):
        ancestor = ancestor.getparent()
    if ancestor is not None:
        try:
            paragraph_index = paragraph_elements.index(ancestor)
        except ValueError:
            paragraph_index = -1
    return (text, paragraph_index)


def _text_between(
    body: etree._Element,
    start: etree._Element,
    end: etree._Element,
) -> str:
    """Concatenate ``<w:t>`` text between ``start`` and ``end`` in document order.

    Walks every descendant of ``body`` exactly once; matches text content
    encountered after ``start`` and before ``end``.
    """
    collecting = False
    parts: list[str] = []
    t_tag = qn("w:t")
    for elem in body.iter():
        if elem is start:
            collecting = True
            continue
        if elem is end:
            break
        if collecting and elem.tag == t_tag and elem.text:
            parts.append(elem.text)
    return "".join(parts)


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


__all__ = ["AnchoredComment", "read_comments"]
