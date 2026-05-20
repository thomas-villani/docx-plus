"""Table of Contents — ``TOC`` complex field.

A TOC in Word is a single complex field whose instruction text
encodes which heading levels to include, whether entries are
hyperlinks, and a handful of other display switches. Word populates
the visible body of the TOC on next open (or when the user presses
F9) — see ``docx_plus.fields.mark_fields_dirty``.

This module imports only from ``docx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import TYPE_CHECKING

from docx_plus.core.oxml import build_complex_field
from docx_plus.publishing._validate import (
    validate_additional_styles,
    validate_outline_levels,
)

if TYPE_CHECKING:
    from docx.text.paragraph import Paragraph
    from lxml import etree


def add_toc(
    paragraph: Paragraph,
    *,
    levels: tuple[int, int] = (1, 3),
    hyperlink: bool = True,
    page_numbers: bool = True,
    additional_styles: Sequence[tuple[str, int]] | None = None,
) -> etree._Element:
    r"""Append a Table of Contents complex field to ``paragraph``.

    Emits a ``TOC`` field whose instruction text matches what Word's
    "Insert → Table of Contents" UI produces. The field has no visible
    result until Word populates it; for that to happen on next open,
    call :func:`docx_plus.fields.mark_fields_dirty` before saving.

    Args:
        paragraph: A python-docx :class:`~docx.text.paragraph.Paragraph`
            where the field is appended. Typically a fresh empty
            paragraph at the position the TOC should occupy.
        levels: Inclusive ``(lo, hi)`` outline-level range for entries.
            Default ``(1, 3)`` matches Word's default of including
            Heading1–Heading3. Word's outline runs 1..9; ``lo`` must
            be ≤ ``hi``.
        hyperlink: When ``True`` (default), entries become clickable
            hyperlinks via the ``\h`` switch.
        page_numbers: When ``False``, suppresses page numbers via the
            ``\n`` switch (useful for web-style TOCs).
        additional_styles: Optional sequence of ``(style_name, level)``
            pairs that the TOC should pick up *in addition* to the
            implicit Heading1..Heading<hi> set selected by ``levels``.
            Plumbs to the ``\t`` switch per ECMA-376 17.16.5.61.
            Style names must contain no commas or double quotes;
            levels must be in 1..9.

    Returns:
        The ``<w:r>`` element wrapping the field's ``begin`` ``fldChar``.

    Raises:
        ValueError: If ``levels`` is not a two-int tuple in the 1..9
            range or is reversed; if any ``additional_styles`` entry is
            malformed; if a style name would terminate the ``\t``
            switch (issues.md H11, H12, H13).

    Example:
        >>> from docx import Document
        >>> from docx_plus.publishing import add_toc
        >>> from docx_plus.fields import mark_fields_dirty
        >>> doc = Document()
        >>> add_toc(doc.add_paragraph(), levels=(1, 2),
        ...         additional_styles=[("Caption", 4)])
        >>> mark_fields_dirty(doc)
    """
    lo, hi = validate_outline_levels(levels)
    extras = validate_additional_styles(additional_styles)

    instruction = f' TOC \\o "{lo}-{hi}"'
    if hyperlink:
        instruction += " \\h"
    instruction += " \\z \\u"
    if not page_numbers:
        instruction += " \\n"
    if extras:
        joined = ",".join(f"{name},{level}" for name, level in extras)
        instruction += f' \\t "{joined}"'
    instruction += " "
    return build_complex_field(paragraph._p, instruction, "")


__all__ = ["add_toc"]
