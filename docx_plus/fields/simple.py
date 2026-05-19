"""Insert OOXML complex fields (PAGE, DATE, generic) into paragraphs.

Word fields use the *complex field* syntax: a sequence of runs containing
``w:fldChar`` markers (``begin``/``separate``/``end``) bracketing the field
instruction text (``w:instrText``) and the result text (``w:t``). This
module emits that five-run sequence and appends it to a paragraph.

Word recalculates field results on open only if ``w:updateFields`` is set in
``settings.xml`` — see :func:`docx_plus.fields.update.mark_fields_dirty`.
Initial text supplied here is what Word shows *before* it recalculates, so
the value is meaningful for offline viewers (e.g. ``"1"`` for a PAGE field).

This module imports only from ``docx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Literal

from lxml import etree

from docx_plus.core.oxml import build_complex_field

if TYPE_CHECKING:
    from docx.text.paragraph import Paragraph


PageFieldName = Literal["PAGE", "NUMPAGES", "SECTIONPAGES"]


def add_page_number_field(
    paragraph: Paragraph,
    *,
    field: PageFieldName = "PAGE",
    format: str | None = None,
) -> etree._Element:
    r"""Append a page-number complex field to ``paragraph``.

    Args:
        paragraph: A python-docx :class:`~docx.text.paragraph.Paragraph`. The
            field runs are appended after the paragraph's existing runs.
        field: Which page-number variant. ``"PAGE"`` (current page),
            ``"NUMPAGES"`` (total pages), or ``"SECTIONPAGES"`` (pages in the
            current section).
        format: Optional field switches appended to the instruction. Example:
            ``r"\* ARABIC"`` forces Arabic numerals, ``r"\* ROMAN"`` Roman.
            See ECMA-376 17.16 for the switch syntax.

    Returns:
        The begin ``w:r`` run that marks the start of the field.

    Example:
        >>> from docx import Document
        >>> from docx_plus.fields import add_page_number_field
        >>> doc = Document()
        >>> p = doc.add_paragraph("Page ")
        >>> _ = add_page_number_field(p)
    """
    if format is None:
        instruction = f" {field} "
    else:
        instruction = f" {field} {format} "
    return build_complex_field(paragraph._p, instruction, "1")


def add_date_field(
    paragraph: Paragraph,
    *,
    format: str = "MMMM d, yyyy",
    auto_update: bool = True,
) -> etree._Element:
    """Append a date complex field to ``paragraph``.

    Args:
        paragraph: A python-docx :class:`~docx.text.paragraph.Paragraph`.
        format: A Word date-format string. Common values:
            ``"MMMM d, yyyy"`` (default, e.g. *May 19, 2026*),
            ``"M/d/yyyy"`` (numeric short), ``"dddd, MMMM d, yyyy"`` (long
            with weekday).
        auto_update: ``True`` (default) emits a ``DATE`` field that Word
            recalculates on every open. ``False`` emits a ``CREATEDATE`` field
            that freezes the document's creation date.

    Returns:
        The begin ``w:r`` run that marks the start of the field.

    Example:
        >>> from docx import Document
        >>> from docx_plus.fields import add_date_field
        >>> doc = Document()
        >>> p = doc.add_paragraph("Today: ")
        >>> _ = add_date_field(p, format="M/d/yyyy")
    """
    keyword = "DATE" if auto_update else "CREATEDATE"
    instruction = f' {keyword} \\@ "{format}" '
    return build_complex_field(paragraph._p, instruction, "")


def add_field(
    paragraph: Paragraph,
    *,
    instruction: str,
    initial_text: str = "",
) -> etree._Element:
    r"""Append a generic complex field to ``paragraph``.

    Use this for fields without a dedicated helper (``TOC``, ``REF``,
    ``HYPERLINK``, ``MERGEFIELD``, etc.). The ``instruction`` is wrapped in
    leading/trailing spaces if you don't supply them, since Word's field
    parser requires them.

    Args:
        paragraph: A python-docx :class:`~docx.text.paragraph.Paragraph`.
        instruction: The raw field instruction text without the surrounding
            ``{ }`` braces Word shows in its UI. Example: ``'REF Bookmark1'``
            or ``'TOC \o "1-3" \h'``.
        initial_text: Optional placeholder shown before Word recalculates.

    Returns:
        The begin ``w:r`` run that marks the start of the field.

    Example:
        >>> from docx import Document
        >>> from docx_plus.fields import add_field
        >>> doc = Document()
        >>> p = doc.add_paragraph()
        >>> _ = add_field(p, instruction='TOC \\o "1-3" \\h', initial_text="(TOC)")
    """
    wrapped = f" {instruction.strip()} "
    return build_complex_field(paragraph._p, wrapped, initial_text)


__all__ = [
    "PageFieldName",
    "add_date_field",
    "add_field",
    "add_page_number_field",
]
