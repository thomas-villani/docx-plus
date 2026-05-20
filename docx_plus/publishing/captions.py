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

if TYPE_CHECKING:
    from docx.text.paragraph import Paragraph
    from lxml import etree


def add_caption(
    paragraph: Paragraph,
    label: str,
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
            trailing whitespace (e.g. ``"Figure "``). Empty string
            suppresses the label run entirely.
        caption_type: The ``SEQ`` field's name. Items sharing this name
            are numbered together (so all ``"Figure"`` captions number
            ``1, 2, 3, …``, independent of all ``"Table"`` captions).
            Must match the ``\c`` switch on any downstream Table of
            Figures.
        numbering: Word numbering format passed to ``\* <numbering>``.
            Common values: ``"ARABIC"`` (default — ``1, 2, 3, …``),
            ``"ROMAN"`` (``i, ii, iii, …``), ``"ALPHABETIC"``
            (``A, B, C, …``).

    Returns:
        The ``<w:r>`` element wrapping the field's ``begin`` ``fldChar``.

    Example:
        >>> from docx import Document
        >>> from docx_plus.publishing import add_caption
        >>> doc = Document()
        >>> p = doc.add_paragraph()
        >>> add_caption(p, "Figure ", caption_type="Figure")
        >>> p.add_run(": Architecture overview")  # caption body text
        <docx.text.run.Run object at 0x...>
    """
    if label:
        label_run = el("w:r")
        label_t = sub(label_run, "w:t", **{"xml:space": "preserve"})
        label_t.text = label
        paragraph._p.append(label_run)

    instruction = f" SEQ {caption_type} \\* {numbering} "
    return build_complex_field(paragraph._p, instruction, "1")


__all__ = ["add_caption"]
