"""The ``CT_Border`` shape, shared by page, table, and cell borders.

ECMA-376 uses one complex type for every border edge in the format: the
four sides of ``<w:pgBorders>`` (17.6.10), of ``<w:tblBorders>``
(17.4.39), and of ``<w:tcBorders>`` (17.4.67), plus the inside and
diagonal edges tables add. All of them carry the same four attributes —
style, thickness, color, and offset — so the dataclass and its
serializer live here rather than in whichever capability package needed
them first.

``layout/borders.py`` shipped :class:`Border` in v0.2 and still
re-exports it, so ``docx_plus.layout.Border`` keeps working.

This module depends on nothing above ``core`` (SPEC §9.1).
"""

from __future__ import annotations

import re
from dataclasses import dataclass

# ECMA-376 17.18.79 ST_HexColor: "auto" or six hex digits ("RRGGBB").
_HEX_COLOR_RE = re.compile(r"^(auto|[0-9A-Fa-f]{6})$")

#: ECMA-376 17.18.2 ST_Border caps ``w:sz`` at 96 eighths of a point.
_MAX_BORDER_SIZE = 96

#: The subset of ECMA-376 17.18.2 ``ST_Border`` that Word's own UI emits.
#:
#: The full enumeration runs to 200+ values, most of them clip-art
#: borders ("w:val=cakeSlice"). Validating against the whole list would
#: mean vendoring it and keeping it current for no practical gain, so
#: this checks the artistic-border *shape* instead: anything outside this
#: set is accepted as long as it is a plain identifier, which catches the
#: real mistake (a CSS-ism like ``"1px solid"`` or ``"#000"``) without
#: rejecting a legitimate exotic value.
_COMMON_BORDER_STYLES = frozenset(
    {
        "nil",
        "none",
        "single",
        "thick",
        "double",
        "dotted",
        "dashed",
        "dotDash",
        "dotDotDash",
        "triple",
        "wave",
        "doubleWave",
        "dashSmallGap",
        "dashDotStroked",
        "threeDEmboss",
        "threeDEngrave",
        "outset",
        "inset",
    }
)

_BORDER_STYLE_RE = re.compile(r"^[A-Za-z][A-Za-z0-9]*$")


@dataclass(frozen=True)
class Border:
    """One border edge.

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
        space: Gap between the reference edge and the border, in
            **points**. ECMA-376 caps this at 31.

            The default ``24`` — 1/3 inch — is a *page*-border default:
            it is what Word's UI emits for "Whole document, Box, Default
            settings" paired with ``offset_from="page"``. Word writes
            ``w:space="0"`` on table and cell borders, so
            :mod:`docx_plus.tables` overrides it rather than inheriting
            this default.

    Raises:
        ValueError: If ``color`` is not ``"auto"`` or a six-hex-digit
            ``"RRGGBB"`` string, if ``size`` is outside ``[0, 96]``, or
            if ``style`` is not a bare identifier.
    """

    style: str = "single"
    size: int = 4
    color: str = "auto"
    space: int = 24

    def __post_init__(self) -> None:
        """Validate the fields against their ECMA-376 simple types."""
        if not _HEX_COLOR_RE.match(self.color):
            raise ValueError(
                "Border.color must be 'auto' or a six-hex-digit 'RRGGBB' string; "
                f"got {self.color!r}"
            )
        if not 0 <= self.size <= _MAX_BORDER_SIZE:
            raise ValueError(
                f"Border.size is in eighths of a point and ECMA-376 caps it at "
                f"{_MAX_BORDER_SIZE}; got {self.size!r}"
            )
        if not _BORDER_STYLE_RE.match(self.style):
            raise ValueError(
                "Border.style must be an ECMA-376 17.18.2 ST_Border name such as "
                f"{'/'.join(sorted(_COMMON_BORDER_STYLES)[:4])}; got {self.style!r}"
            )


def border_attrs(border: Border) -> dict[str, str]:
    """Serialize ``border`` to the ``CT_Border`` attribute mapping.

    Every border edge in the format takes the same four attributes, so
    every writer — page, table, cell — goes through this.

    Args:
        border: The border to serialize.

    Returns:
        A ``{"w:val": ..., "w:sz": ..., "w:color": ..., "w:space": ...}``
        mapping ready to splat into :func:`docx_plus.core.oxml.el`.
    """
    return {
        "w:val": border.style,
        "w:sz": str(border.size),
        "w:color": border.color,
        "w:space": str(border.space),
    }


__all__ = ["Border", "border_attrs"]
