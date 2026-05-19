"""Cross-references to bookmarks via ``REF`` / ``PAGEREF`` fields.

``REF bookmark_name`` inserts the text Word reads from the bookmark's
range; ``PAGEREF bookmark_name`` inserts the page number Word renders
for the bookmark. Both are complex fields built on top of the same
plumbing :mod:`docx_plus.fields` uses for page numbers and dates.

This module imports only from ``docx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Literal

from lxml import etree

from docx_plus.core.oxml import build_complex_field

if TYPE_CHECKING:
    from docx.text.paragraph import Paragraph

CrossReferenceKind = Literal["text", "page"]


def add_cross_reference(
    paragraph: Paragraph,
    *,
    bookmark: str,
    kind: CrossReferenceKind = "text",
    hyperlink: bool = True,
) -> etree._Element:
    r"""Append a cross-reference complex field to ``paragraph``.

    Args:
        paragraph: A python-docx :class:`~docx.text.paragraph.Paragraph`
            where the cross-reference field is appended after any
            existing runs.
        bookmark: The target bookmark's ``w:name`` attribute. Must match
            an existing bookmark for Word to resolve the field;
            unresolved cross-references render as ``"Error! Reference
            source not found."``. Bookmark names must match Word's
            rules — see :func:`docx_plus.bookmarks.add_bookmark`.
        kind: ``"text"`` (default) inserts a ``REF`` field that resolves
            to the bookmark's text content; ``"page"`` inserts a
            ``PAGEREF`` field that resolves to the page number where
            the bookmark sits.
        hyperlink: ``True`` (default) appends ``\h`` to the field
            instruction so Word makes the cross-reference a clickable
            link to the bookmark.

    Returns:
        The begin ``w:r`` run that marks the start of the field, same
        contract as :func:`docx_plus.fields.add_page_number_field`.

    Example:
        >>> from docx import Document
        >>> from docx_plus.bookmarks import add_bookmark, add_cross_reference
        >>> doc = Document()
        >>> p1 = doc.add_paragraph("Section 1")
        >>> add_bookmark(p1, "sec_1")
        >>> p2 = doc.add_paragraph("See ")
        >>> add_cross_reference(p2, bookmark="sec_1", kind="text")

    Notes:
        Fields are cached: Word displays the previously-computed result
        until ``w:updateFields="true"`` triggers recalculation on open.
        Pair calls to :func:`add_cross_reference` with
        :func:`docx_plus.fields.mark_fields_dirty` so the new
        cross-references resolve on first open.
    """
    keyword = "REF" if kind == "text" else "PAGEREF"
    flag = " \\h" if hyperlink else ""
    instruction = f" {keyword} {bookmark}{flag} "
    return build_complex_field(paragraph._p, instruction, "")


__all__ = ["CrossReferenceKind", "add_cross_reference"]
