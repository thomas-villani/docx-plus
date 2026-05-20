"""Multi-column page layout for a section (``<w:cols>``).

python-docx exposes orientation, margins, page size, and header/footer
settings on :class:`docx.section.Section`, but does not provide an
abstraction for the ``<w:cols>`` child of ``<w:sectPr>``. This module
fills that gap with a single :func:`set_columns` helper.

ECMA-376 §17.6.4: ``cols`` controls the multi-column layout of a
section. ``w:num`` is the column count, ``w:space`` is the spacing
between columns (twips), ``w:sep`` enables vertical separator lines,
and child ``<w:col w:w=... w:space=.../>`` elements set unequal widths.

This module imports only from ``docx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import TYPE_CHECKING

from docx_plus.core.ns import qn
from docx_plus.core.oxml import el, insert_before_first_anchor, remove, sub

if TYPE_CHECKING:
    from docx.section import Section


# Schema siblings later than `w:cols` per ECMA-376 17.6.17 CT_SectPr.
_LATER_SIBLINGS: tuple[str, ...] = (
    "w:formProt",
    "w:vAlign",
    "w:noEndnote",
    "w:titlePg",
    "w:textDirection",
    "w:bidi",
    "w:rtlGutter",
    "w:docGrid",
    "w:printerSettings",
    "w:sectPrChange",
)


def set_columns(
    section: Section,
    num: int,
    *,
    space: int = 720,
    separator: bool = False,
    widths: Sequence[int] | None = None,
) -> None:
    """Configure multi-column layout for ``section``.

    Idempotent: a call replaces any existing ``<w:cols>`` rather than
    stacking elements. Schema-strict — the element lands in its
    ECMA-376 17.6.17 slot regardless of which other ``<w:sectPr>``
    children (e.g. ``<w:docGrid>``) are already present.

    Args:
        section: A python-docx :class:`~docx.section.Section`. The
            mutated element is the section's underlying
            ``<w:sectPr>`` (``section._sectPr``).
        num: Number of columns. Must be ``>= 1``.
        space: Spacing between columns in twips (1/1440 inch). Default
            720 twips = 0.5 inch. Ignored when ``widths`` provides
            per-column spacing.
        separator: When ``True``, emit ``w:sep="1"`` so Word renders a
            vertical line between columns. Default ``False``.
        widths: Optional unequal-width spec, in twips. Length must equal
            ``num``. When supplied, Word reads per-column widths from
            child ``<w:col>`` elements and ignores the parent ``w:space``;
            the child elements share ``space`` between adjacent columns
            (last column omits trailing space).

    Raises:
        ValueError: If ``num < 1`` or ``len(widths) != num``.

    Example:
        >>> from docx import Document
        >>> from docx_plus.layout import set_columns
        >>> doc = Document()
        >>> set_columns(doc.sections[0], 2, space=720, separator=True)
    """
    if num < 1:
        raise ValueError(f"set_columns requires num >= 1, got {num}")
    if widths is not None and len(widths) != num:
        raise ValueError(
            f"widths has {len(widths)} entries but num={num} columns requested"
        )

    sect_pr = section._sectPr
    existing = sect_pr.find(qn("w:cols"))
    if existing is not None:
        remove(existing)

    attrs = {"w:num": str(num), "w:space": str(space)}
    if separator:
        attrs["w:sep"] = "1"
    if widths is not None:
        # Per-column widths supersede the simple `space` attr; Word still
        # requires `w:equalWidth="0"` for the per-col branch to be honored.
        attrs["w:equalWidth"] = "0"

    cols = el("w:cols", **attrs)
    if widths is not None:
        for idx, w in enumerate(widths):
            col_attrs = {"w:w": str(w)}
            if idx < num - 1:
                col_attrs["w:space"] = str(space)
            sub(cols, "w:col", **col_attrs)

    insert_before_first_anchor(sect_pr, cols, _LATER_SIBLINGS)


__all__ = ["set_columns"]
