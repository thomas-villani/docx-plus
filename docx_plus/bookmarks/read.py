"""Read every bookmark from a document.

Returns one :class:`BookmarkInfo` per ``<w:bookmarkStart>`` paired with
its matching ``<w:bookmarkEnd>``. The anchored text is what a ``REF
bookmark_name`` field would resolve to.

This module imports only from ``docx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from lxml import etree

from docx_plus.core.ns import qn
from docx_plus.core.oxml import xpath

if TYPE_CHECKING:
    from docx.document import Document


@dataclass(frozen=True)
class BookmarkInfo:
    """A bookmark with the text it anchors to.

    Attributes:
        bookmark_id: The ``w:id`` value.
        name: The ``w:name`` attribute. Cross-references key off the
            name, not the id.
        anchored_text: The text between ``bookmarkStart`` and
            ``bookmarkEnd``. Empty for unclosed or empty-range
            bookmarks.
        paragraph_index: Zero-based index (within
            ``doc.paragraphs``) of the paragraph that contains the
            ``bookmarkStart`` marker. ``-1`` if the bookmark sits
            outside any paragraph (rare; structurally invalid).
    """

    bookmark_id: int
    name: str
    anchored_text: str
    paragraph_index: int


def read_bookmarks(doc: Document) -> list[BookmarkInfo]:
    """Return every bookmark in ``doc`` paired with the text it anchors to.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to scan.

    Returns:
        One :class:`BookmarkInfo` per bookmark, in document order.
    """
    body = doc.element.body
    paragraph_elements = list(xpath(body, ".//w:p"))

    # Build an id → end-element map so we don't re-scan for each start.
    ends_by_id: dict[str, etree._Element] = {}
    for end in xpath(body, ".//w:bookmarkEnd"):
        bid = end.get(qn("w:id"))
        if bid is not None:
            ends_by_id[bid] = end

    result: list[BookmarkInfo] = []
    for start in xpath(body, ".//w:bookmarkStart"):
        bid_raw = start.get(qn("w:id"))
        name = start.get(qn("w:name")) or ""
        if bid_raw is None:
            continue
        try:
            bid = int(bid_raw)
        except ValueError:
            continue

        end = ends_by_id.get(bid_raw)
        anchored_text = _text_between(body, start, end) if end is not None else ""

        paragraph_index = -1
        ancestor = start.getparent()
        while ancestor is not None and ancestor.tag != qn("w:p"):
            ancestor = ancestor.getparent()
        if ancestor is not None:
            try:
                paragraph_index = paragraph_elements.index(ancestor)
            except ValueError:
                paragraph_index = -1

        result.append(
            BookmarkInfo(
                bookmark_id=bid,
                name=name,
                anchored_text=anchored_text,
                paragraph_index=paragraph_index,
            )
        )
    return result


def _text_between(
    body: etree._Element,
    start: etree._Element,
    end: etree._Element,
) -> str:
    """Concatenate ``<w:t>`` text between ``start`` and ``end`` in document order."""
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


__all__ = ["BookmarkInfo", "read_bookmarks"]
