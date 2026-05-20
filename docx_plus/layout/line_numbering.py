"""Section line numbering (``<w:lnNumType>``).

python-docx does not abstract ``<w:lnNumType>``, the section-scoped
control for printing line numbers in the margin (the legal / contract
use case). This module fills the gap with a single
:func:`set_line_numbering` helper.

ECMA-376 §17.6.8: ``lnNumType`` is an attributes-only element with
``countBy`` (interval between printed numbers), ``start`` (the first
line number for the section), ``distance`` (twips between text and
number), and ``restart`` controlling whether numbering restarts on
each page / each new section / never (``continuous``).

This module imports only from ``docx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Literal

from docx_plus.core.ns import qn
from docx_plus.core.oxml import el, insert_before_first_anchor, remove

if TYPE_CHECKING:
    from docx.section import Section


LineNumberRestart = Literal["newPage", "newSection", "continuous"]


# Schema siblings later than `w:lnNumType` per ECMA-376 17.6.17 CT_SectPr.
_LATER_SIBLINGS: tuple[str, ...] = (
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


def set_line_numbering(
    section: Section,
    *,
    count_by: int = 1,
    restart: LineNumberRestart = "newPage",
    start: int = 1,
    distance: int | None = None,
) -> None:
    """Enable line numbering for ``section``.

    Idempotent: replaces any existing ``<w:lnNumType>`` rather than
    stacking elements. Schema-strict — the element lands in its
    ECMA-376 17.6.17 slot regardless of which other ``<w:sectPr>``
    children are present.

    Args:
        section: A python-docx :class:`~docx.section.Section`.
        count_by: Interval at which line numbers are printed. ``1``
            (default) prints every line, ``5`` prints every fifth, etc.
            Must be ``>= 1``.
        restart: When numbering restarts.

            - ``"newPage"`` (default) — restart at 1 on each page
            - ``"newSection"`` — restart at 1 at each section break
            - ``"continuous"`` — never restart; run through the document
        start: First line number for this section. Must be ``>= 1``.
            Word still applies ``count_by`` from this starting value.
        distance: Twips between the text and the printed number. ``None``
            (default) omits the attribute so Word uses its built-in
            default (varies by version; ~360 twips / 0.25 inch is
            typical).

    Raises:
        ValueError: If ``count_by < 1``, ``start < 1``, ``distance`` is
            negative, or ``restart`` is not one of the three allowed values.

    Example:
        >>> from docx import Document
        >>> from docx_plus.layout import set_line_numbering
        >>> doc = Document()
        >>> set_line_numbering(doc.sections[0], count_by=5, restart="continuous")
    """
    if count_by < 1:
        raise ValueError(f"set_line_numbering requires count_by >= 1, got {count_by}")
    if start < 1:
        raise ValueError(f"set_line_numbering requires start >= 1, got {start}")
    if restart not in ("newPage", "newSection", "continuous"):
        raise ValueError(
            f'restart must be "newPage", "newSection", or "continuous"; got "{restart}"'
        )
    if distance is not None and distance < 0:
        raise ValueError(f"set_line_numbering requires distance >= 0, got {distance}")

    sect_pr = section._sectPr
    existing = sect_pr.find(qn("w:lnNumType"))
    if existing is not None:
        remove(existing)

    attrs = {
        "w:countBy": str(count_by),
        "w:start": str(start),
        "w:restart": restart,
    }
    if distance is not None:
        attrs["w:distance"] = str(distance)

    ln = el("w:lnNumType", **attrs)
    insert_before_first_anchor(sect_pr, ln, _LATER_SIBLINGS)


__all__ = ["LineNumberRestart", "set_line_numbering"]
