r"""Figure / table captions — leading text + ``SEQ`` complex field.

A Word caption is a paragraph that opens with a label run (e.g.
``"Figure "``) followed by a ``SEQ`` complex field that auto-numbers
items of the same caption type. The Table of Figures (see
``docx_plus.publishing.figures``) picks up captions whose ``SEQ`` name
matches its ``\\c`` switch.

This module imports only from ``docx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx_plus.core.oxml import build_complex_field, el, sub
from docx_plus.publishing._validate import (
    validate_numbering_picture,
    validate_seq_identifier,
)

if TYPE_CHECKING:
    from docx.text.paragraph import Paragraph
    from lxml import etree


def add_caption(
    paragraph: Paragraph,
    label: str | None = None,
    *,
    caption_type: str = "Figure",
    numbering: str = "ARABIC",
) -> etree._Element:
    r"""Append a caption (label run + auto-numbered ``SEQ`` field).

    The label is emitted as a literal text run; the number is a
    ``SEQ`` complex field that Word re-numbers on open. After all
    captions are inserted, call
    :func:`docx_plus.fields.mark_fields_dirty` so Word recalculates
    the SEQ values.

    Args:
        paragraph: A python-docx :class:`~docx.text.paragraph.Paragraph`
            where the caption is appended. Typically a fresh paragraph
            beneath the figure / table being captioned.
        label: Leading text shown before the number, including any
            trailing whitespace. When omitted (``None``, the default),
            uses ``f"{caption_type} "`` — the common case. Pass an
            empty string ``""`` to suppress the label run entirely
            (e.g. when the surrounding paragraph already supplies it).
        caption_type: The ``SEQ`` field's name. Items sharing this name
            are numbered together (so all ``"Figure"`` captions number
            ``1, 2, 3, …``, independent of all ``"Table"`` captions).
            Must match the ``\c`` switch on any downstream Table of
            Figures, and must conform to the SEQ identifier rule
            (ASCII letter/underscore start, then letters/digits/
            underscores).
        numbering: Word numbering format token for the ``\* <picture>``
            switch. Common values: ``"ARABIC"`` (default — ``1, 2, 3,
            …``), ``"ROMAN"`` (``I, II, III, …``), ``"roman"``
            (``i, ii, iii, …``), ``"ALPHABETIC"`` (``A, B, C, …``).
            See ECMA-376 17.16.4.1 for the full token list.

    Returns:
        The ``<w:r>`` element wrapping the field's ``begin`` ``fldChar``.

    Raises:
        ValueError: If ``caption_type`` is empty or violates the SEQ
            identifier rule, or if ``numbering`` is not a recognised
            format token (issues.md H11, M16).

    Note:
        The caption's paragraph is *not* automatically restyled to
        Word's built-in ``Caption`` paragraph style. Apply it yourself
        if you want the conventional italic-grey rendering:
        ``paragraph.style = doc.styles["Caption"]``.

    Example:
        >>> from docx import Document
        >>> from docx_plus.publishing import add_caption
        >>> doc = Document()
        >>> p = doc.add_paragraph()
        >>> add_caption(p, caption_type="Figure")  # label defaults to "Figure "
        >>> p.add_run(": Architecture overview")
        <docx.text.run.Run object at 0x...>
    """
    validate_seq_identifier(caption_type, arg_name="caption_type")
    validate_numbering_picture(numbering)

    if label is None:
        label = f"{caption_type} "

    if label:
        label_run = el("w:r")
        label_t = sub(label_run, "w:t", **{"xml:space": "preserve"})
        label_t.text = label
        paragraph._p.append(label_run)

    instruction = f" SEQ {caption_type} \\* {numbering} "
    return build_complex_field(paragraph._p, instruction, "1")


__all__ = ["add_caption"]
