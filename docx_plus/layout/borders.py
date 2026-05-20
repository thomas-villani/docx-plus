"""Page borders (``<w:pgBorders>``).

python-docx does not abstract ``<w:pgBorders>`` — the section-scoped
control for the decorative box around a page that formal documents
(certificates, awards, contract title pages) frequently want. This
module fills the gap with a :class:`Border` dataclass and a single
:func:`set_page_borders` helper.

ECMA-376 §17.6.10: ``pgBorders`` is a container element whose four
optional children (``top``, ``left``, ``bottom``, ``right`` — in
schema order) each declare their style (``w:val``), thickness in
eighths of a point (``w:sz``), color (``w:color``), and the gap from
the reference edge (``w:space``, in points). The container also takes
an ``offsetFrom`` attribute that selects whether ``w:space`` is
measured from the page edge or the body text.

This module imports only from ``docx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING, Literal

from docx_plus.core.ns import qn
from docx_plus.core.oxml import el, insert_before_first_anchor, remove, sub

if TYPE_CHECKING:
    from docx.section import Section


# ECMA-376 17.18.79 ST_HexColor: "auto" or six hex digits ("RRGGBB").
_HEX_COLOR_RE = re.compile(r"^(auto|[0-9A-Fa-f]{6})$")


# Schema siblings later than `w:pgBorders` per ECMA-376 17.6.17 CT_SectPr.
_LATER_SIBLINGS: tuple[str, ...] = (
    "w:lnNumType",
    "w:pgNumType",
    "w:cols",
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

OffsetFrom = Literal["page", "text"]


@dataclass(frozen=True)
class Border:
    """One side of a page border.

    Attributes:
        style: ECMA-376 17.18.2 border style name. Common values:
            ``"single"`` (default), ``"double"``, ``"thick"``,
            ``"dashed"``, ``"dotted"``, ``"wave"``, ``"none"``. The
            full enumeration has 200+ entries — see the spec.
        size: Border thickness in **eighths of a point** (so ``4`` is
            0.5 pt and ``8`` is 1 pt). ECMA-376 caps this at 96.
        color: ``"RRGGBB"`` hex or ``"auto"`` (default) to let Word pick
            a sensible contrast. Validated at construction against
            ECMA-376 17.18.79 ``ST_HexColor`` — ``"red"``, ``"#FF0000"``,
            or a 3-digit shorthand raise :class:`ValueError`.
        space: Gap between the page edge (or text — see
            :func:`set_page_borders`'s ``offset_from``) and the border,
            in **points**. ECMA-376 17.6.10 caps this at 31. ``24``
            (default) — 1/3 inch — matches what Word's UI emits for
            ``"Whole document, Box, Default settings"`` paired with the
            default ``offset_from="page"``.

    Raises:
        ValueError: If ``color`` is not ``"auto"`` or a six-hex-digit
            ``"RRGGBB"`` string.
    """

    style: str = "single"
    size: int = 4
    color: str = "auto"
    space: int = 24

    def __post_init__(self) -> None:
        """Validate ``color`` against ECMA-376 ``ST_HexColor`` at construction."""
        if not _HEX_COLOR_RE.match(self.color):
            raise ValueError(
                "Border.color must be 'auto' or a six-hex-digit 'RRGGBB' string; "
                f"got {self.color!r}"
            )


def set_page_borders(
    section: Section,
    *,
    top: Border | None = None,
    bottom: Border | None = None,
    left: Border | None = None,
    right: Border | None = None,
    offset_from: OffsetFrom = "page",
) -> None:
    """Configure the page border for ``section``.

    Idempotent: replaces any existing ``<w:pgBorders>``. Passing all
    four sides as ``None`` removes the element instead of writing an
    empty container. Child sides are written in the schema-required
    order ``top → left → bottom → right`` per ECMA-376 17.6.10.

    Args:
        section: A python-docx :class:`~docx.section.Section`.
        top: Border for the top edge, or ``None`` to omit.
        bottom: Border for the bottom edge.
        left: Border for the left edge.
        right: Border for the right edge.
        offset_from: ``"page"`` (default) measures ``Border.space`` from
            the page edge — what Word's UI emits and what callers
            usually want for a decorative frame. ``"text"`` measures
            from the body text edge — the spec's *implicit* default
            when the attribute is omitted, producing a tight inner box.

    Example:
        >>> from docx import Document
        >>> from docx_plus.layout import Border, set_page_borders
        >>> doc = Document()
        >>> rule = Border(style="single", size=8, color="2F5496")
        >>> set_page_borders(doc.sections[0], top=rule, bottom=rule,
        ...                  left=rule, right=rule)
    """
    sect_pr = section._sectPr
    existing = sect_pr.find(qn("w:pgBorders"))
    if existing is not None:
        remove(existing)

    if top is None and bottom is None and left is None and right is None:
        return

    borders_el = el("w:pgBorders", **{"w:offsetFrom": offset_from})
    # Schema-required order per ECMA-376 17.6.10 CT_PageBorders.
    for tag, border in (
        ("w:top", top),
        ("w:left", left),
        ("w:bottom", bottom),
        ("w:right", right),
    ):
        if border is None:
            continue
        sub(
            borders_el,
            tag,
            **{
                "w:val": border.style,
                "w:sz": str(border.size),
                "w:color": border.color,
                "w:space": str(border.space),
            },
        )

    insert_before_first_anchor(sect_pr, borders_el, _LATER_SIBLINGS)


__all__ = ["Border", "OffsetFrom", "set_page_borders"]
