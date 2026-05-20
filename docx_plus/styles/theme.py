"""Read-only theme color resolution.

WordprocessingML references theme colors symbolically (``themeColor="accent1"``)
with optional ``themeTint``/``themeShade`` modifiers. The actual RGB values
live in ``word/theme/theme1.xml`` under ``a:clrScheme``. This module reads that
scheme, translates Word's ``ST_ThemeColor`` names to DrawingML scheme keys
(ECMA-376 17.18.97), and applies the tint/shade/lumMod/lumOff transforms
defined in ECMA-376 17.18.40.

Failures here are recoverable: a missing or malformed theme part is reported
by :func:`load_theme` returning ``None`` (or a partially-populated scheme),
not by raising. Callers — primarily the cascade resolver — fold that into a
``partial=True`` flag on the resolved formatting. SPEC §4 "Theme references".

The same scheme also exposes the theme's *fonts* (``a:fontScheme``):
:func:`resolve_theme_font` maps a WordprocessingML font-theme token
(``w:asciiTheme="minorHAnsi"`` etc.) to the concrete typeface the theme
defines (``"Calibri"``), so the cascade can report a real font name rather
than the bare token.

The ``w:color`` cascade element (ECMA-376 CT_Color) carries only
``themeTint`` / ``themeShade``, so :func:`resolve_theme_color` applies just
those two transforms. :func:`apply_lum_mod` / :func:`apply_lum_off`
implement the DrawingML ``lumMod`` / ``lumOff`` transforms for callers that
read theme colors *referenced from DrawingML* (shape fills, ``w14`` text
effects), where those transforms do appear — they are deliberately not part
of the ``w:color`` resolution path because that element cannot carry them.

The module is read-only; writing themes is a v0.2 non-goal (SPEC §1).
"""

from __future__ import annotations

import colorsys
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from lxml import etree

from docx_plus.core import DocxPlusError
from docx_plus.core.ns import A, qn
from docx_plus.core.oxml import xpath

if TYPE_CHECKING:
    from docx.document import Document


_THEME_RELTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"

# Word's ST_ThemeColor name -> DrawingML clrScheme child element name.
# Per ECMA-376 17.18.97: text1/background1/text2/background2 are aliases for
# dk1/lt1/dk2/lt2 respectively.
_THEME_NAME_TO_SCHEME_KEY: dict[str, str] = {
    "dark1": "dk1",
    "light1": "lt1",
    "dark2": "dk2",
    "light2": "lt2",
    "accent1": "accent1",
    "accent2": "accent2",
    "accent3": "accent3",
    "accent4": "accent4",
    "accent5": "accent5",
    "accent6": "accent6",
    "hyperlink": "hlink",
    "followedHyperlink": "folHlink",
    "text1": "dk1",
    "text2": "dk2",
    "background1": "lt1",
    "background2": "lt2",
}


class ThemeError(DocxPlusError):
    """Raised when theme inputs are structurally invalid in an unrecoverable way.

    Most theme defects (missing part, malformed XML, unknown name) are reported
    via ``None`` returns or ``partial=True`` per SPEC §4. This error is
    reserved for programmer-error cases such as an unparseable hex transform
    byte that would otherwise pass through silently.
    """


@dataclass(frozen=True)
class ThemeColors:
    """Resolved theme color + font scheme.

    Built by :func:`load_theme`. Use :meth:`base` to look up a color by
    Word's ``ST_ThemeColor`` name and :meth:`font` to look up a typeface
    by ``ST_Theme`` font token (both are what appear in WordprocessingML).

    Attributes:
        scheme: DrawingML color key (``"accent1"``, ``"dk1"``, …) ->
            uppercase ``RRGGBB`` hex.
        fonts: ``ST_Theme`` font token (``"minorHAnsi"``,
            ``"majorEastAsia"``, …) -> concrete typeface name. Empty when
            the theme has no ``a:fontScheme``.
    """

    scheme: dict[str, str]
    fonts: dict[str, str] = field(default_factory=dict)

    def base(self, theme_name: str) -> str | None:
        """Return the unmodified hex color for a Word theme color name.

        Args:
            theme_name: A value from ``ST_ThemeColor`` (e.g. ``"accent1"``,
                ``"text1"``).

        Returns:
            Uppercase ``RRGGBB`` hex string, or ``None`` if the name is not a
            recognized theme color or the underlying scheme entry is missing.
        """
        key = _THEME_NAME_TO_SCHEME_KEY.get(theme_name)
        if key is None:
            return None
        return self.scheme.get(key)

    def font(self, token: str) -> str | None:
        """Return the concrete typeface for an ``ST_Theme`` font token.

        Args:
            token: A ``w:asciiTheme`` / ``w:hAnsiTheme`` /
                ``w:eastAsiaTheme`` / ``w:cstheme`` value such as
                ``"minorHAnsi"`` or ``"majorEastAsia"``.

        Returns:
            The typeface name from the theme's ``a:fontScheme`` (e.g.
            ``"Calibri"``), or ``None`` if the token is unknown or the
            scheme entry is empty / missing.
        """
        return self.fonts.get(token)


def load_theme(doc: Document) -> ThemeColors | None:
    """Read ``word/theme/theme1.xml`` and return its color scheme.

    Args:
        doc: A python-docx :class:`~docx.document.Document`.

    Returns:
        A :class:`ThemeColors` describing the document's color scheme, or
        ``None`` if the document has no theme part attached or the theme part
        cannot be parsed at all. A partially-readable scheme is returned as a
        ``ThemeColors`` whose ``.scheme`` dict simply omits the unreadable
        entries — callers can detect partiality via :meth:`ThemeColors.base`
        returning ``None``.
    """
    theme_xml = _read_theme_blob(doc)
    if theme_xml is None:
        return None
    try:
        root = etree.fromstring(theme_xml)
    except etree.XMLSyntaxError:
        return None
    return ThemeColors(scheme=_parse_clr_scheme(root), fonts=_parse_font_scheme(root))


def resolve_theme_font(theme: ThemeColors | None, token: str) -> str | None:
    """Resolve a WordprocessingML font-theme token to a concrete typeface.

    Args:
        theme: The document's theme, or ``None`` if no theme part is
            attached. ``None`` always resolves to ``None``.
        token: An ``ST_Theme`` value (e.g. ``"minorHAnsi"``,
            ``"majorEastAsia"``).

    Returns:
        The typeface name (e.g. ``"Calibri"``), or ``None`` if the theme is
        absent or the token has no entry in the font scheme.
    """
    if theme is None:
        return None
    return theme.font(token)


def resolve_theme_color(
    theme: ThemeColors | None,
    name: str,
    *,
    tint: str | None = None,
    shade: str | None = None,
) -> str | None:
    """Resolve a WordprocessingML theme color reference to ``RRGGBB`` hex.

    Args:
        theme: The document's theme scheme, or ``None`` if no theme part is
            attached. ``None`` always resolves to ``None``.
        name: Word ``ST_ThemeColor`` value (e.g. ``"accent1"``, ``"text1"``,
            ``"none"``).
        tint: Optional ``w:themeTint`` value — a hex byte ``"00"``-``"FF"``.
            Lightens the resolved color toward white.
        shade: Optional ``w:themeShade`` value — a hex byte ``"00"``-``"FF"``.
            Darkens the resolved color toward black.

    Returns:
        Uppercase ``RRGGBB`` hex string, or ``None`` if the name is unknown,
        the theme is absent, or the name is the literal ``"none"``.

    Note:
        WordprocessingML treats ``themeTint`` and ``themeShade`` as mutually
        exclusive in practice, but this function tolerates both being set: the
        shade is applied first, then the tint, matching the order Word uses
        when it encounters the (unusual) combination.
    """
    if theme is None or name == "none":
        return None
    base = theme.base(name)
    if base is None:
        return None
    out = base
    if shade is not None:
        out = apply_theme_shade(out, shade)
    if tint is not None:
        out = apply_theme_tint(out, tint)
    return out


def apply_theme_tint(hex_color: str, tint_byte: str) -> str:
    """Lighten ``hex_color`` toward white by ECMA-376 17.18.40 ``themeTint``.

    Algorithm: convert to HSL, replace ``L`` with ``L * t + (1 - t)`` where
    ``t = int(tint_byte, 16) / 255``. ``tint="FF"`` is a no-op; ``tint="00"``
    forces L to 1 (pure white).

    Args:
        hex_color: Six-character hex color (with or without leading ``#``).
        tint_byte: Hex byte ``"00"``-``"FF"``.

    Returns:
        Uppercase ``RRGGBB`` hex string.
    """
    t = _parse_hex_byte(tint_byte)
    h, lum, s = _rgb_to_hls(hex_color)
    new_l = lum * t + (1 - t)
    return _hls_to_hex(h, new_l, s)


def apply_theme_shade(hex_color: str, shade_byte: str) -> str:
    """Darken ``hex_color`` toward black by ECMA-376 17.18.40 ``themeShade``.

    Algorithm: convert to HSL, replace ``L`` with ``L * s`` where
    ``s = int(shade_byte, 16) / 255``. ``shade="FF"`` is a no-op; ``shade="00"``
    forces L to 0 (pure black).

    Args:
        hex_color: Six-character hex color (with or without leading ``#``).
        shade_byte: Hex byte ``"00"``-``"FF"``.

    Returns:
        Uppercase ``RRGGBB`` hex string.
    """
    s_val = _parse_hex_byte(shade_byte)
    h, lum, sat = _rgb_to_hls(hex_color)
    return _hls_to_hex(h, lum * s_val, sat)


def apply_lum_mod(hex_color: str, lum_mod: int) -> str:
    """Multiply L by ``lum_mod / 100000`` per ECMA-376 17.18.40.

    DrawingML transform values are percent thousandths: ``50000`` means 50%.

    Args:
        hex_color: Six-character hex color.
        lum_mod: Percent thousandths (e.g. ``50000`` for 50%).

    Returns:
        Uppercase ``RRGGBB`` hex string with L clamped to ``[0, 1]``.
    """
    factor = lum_mod / 100000.0
    h, lum, sat = _rgb_to_hls(hex_color)
    return _hls_to_hex(h, lum * factor, sat)


def apply_lum_off(hex_color: str, lum_off: int) -> str:
    """Add ``lum_off / 100000`` to L per ECMA-376 17.18.40.

    DrawingML transform values are percent thousandths: ``80000`` adds 0.8.
    The result is clamped to ``[0, 1]``.

    Args:
        hex_color: Six-character hex color.
        lum_off: Percent thousandths (e.g. ``80000`` for +0.8).

    Returns:
        Uppercase ``RRGGBB`` hex string.
    """
    delta = lum_off / 100000.0
    h, lum, sat = _rgb_to_hls(hex_color)
    return _hls_to_hex(h, lum + delta, sat)


def _read_theme_blob(doc: Document) -> bytes | None:
    document_part = doc.part
    for rel in document_part.rels.values():
        if rel.reltype == _THEME_RELTYPE:
            target = rel.target_part
            blob = getattr(target, "blob", None)
            if isinstance(blob, bytes):
                return blob
    return None


def _parse_clr_scheme(theme_root: etree._Element) -> dict[str, str]:
    out: dict[str, str] = {}
    children = xpath(theme_root, "./a:themeElements/a:clrScheme/*")
    for scheme_child in children:
        if not isinstance(scheme_child, etree._Element):
            continue
        qname = etree.QName(scheme_child.tag)
        if qname.namespace != A:
            continue
        color = _extract_color(scheme_child)
        if color is not None:
            out[qname.localname] = color
    return out


def _parse_font_scheme(theme_root: etree._Element) -> dict[str, str]:
    """Map ECMA-376 ``ST_Theme`` font tokens to concrete typeface names.

    Reads ``a:themeElements/a:fontScheme``. Each of ``majorFont`` /
    ``minorFont`` carries ``a:latin`` / ``a:ea`` / ``a:cs`` typefaces; the
    WordprocessingML font-theme tokens map on top of those — ``*Ascii`` and
    ``*HAnsi`` -> latin, ``*EastAsia`` -> ea, ``*Bidi`` -> cs (ECMA-376
    20.1.4.1.24). Empty typefaces are omitted so an unresolved token surfaces
    as ``None`` rather than an empty string.
    """
    out: dict[str, str] = {}
    scheme_matches = xpath(theme_root, "./a:themeElements/a:fontScheme")
    if not scheme_matches:
        return out
    scheme = scheme_matches[0]
    if not isinstance(scheme, etree._Element):
        return out
    for font_tag, prefix in (("a:majorFont", "major"), ("a:minorFont", "minor")):
        font_el = scheme.find(qn(font_tag))
        if font_el is None:
            continue
        latin = _typeface(font_el.find(qn("a:latin")))
        if latin is not None:
            out[f"{prefix}Ascii"] = latin
            out[f"{prefix}HAnsi"] = latin
        ea = _typeface(font_el.find(qn("a:ea")))
        if ea is not None:
            out[f"{prefix}EastAsia"] = ea
        cs = _typeface(font_el.find(qn("a:cs")))
        if cs is not None:
            out[f"{prefix}Bidi"] = cs
    return out


def _typeface(latin_or_ea_or_cs: etree._Element | None) -> str | None:
    """Return a non-empty ``typeface`` attribute, or ``None``."""
    if latin_or_ea_or_cs is None:
        return None
    typeface = latin_or_ea_or_cs.get("typeface")
    return typeface or None


def _extract_color(scheme_child: etree._Element) -> str | None:
    """Read the RRGGBB hex from a clrScheme child (e.g. ``a:accent1``)."""
    srgb = scheme_child.find(qn("a:srgbClr"))
    if srgb is not None:
        val = srgb.get("val")
        return val.upper() if val else None
    sys_clr = scheme_child.find(qn("a:sysClr"))
    if sys_clr is not None:
        last = sys_clr.get("lastClr")
        return last.upper() if last else None
    return None


def _parse_hex_byte(byte_str: str) -> float:
    try:
        value = int(byte_str, 16)
    except (TypeError, ValueError) as exc:
        raise ThemeError(f"expected hex byte, got {byte_str!r}") from exc
    if not 0 <= value <= 0xFF:
        raise ThemeError(f"hex byte {byte_str!r} out of range")
    return value / 255.0


def _rgb_to_hls(hex_color: str) -> tuple[float, float, float]:
    cleaned = hex_color.lstrip("#")
    if len(cleaned) != 6:
        raise ThemeError(f"expected 6-character hex color, got {hex_color!r}")
    try:
        r = int(cleaned[0:2], 16)
        g = int(cleaned[2:4], 16)
        b = int(cleaned[4:6], 16)
    except ValueError as exc:
        raise ThemeError(f"unparseable hex color {hex_color!r}") from exc
    return colorsys.rgb_to_hls(r / 255.0, g / 255.0, b / 255.0)


def _hls_to_hex(h: float, lum: float, s: float) -> str:
    clamped_l = max(0.0, min(1.0, lum))
    r, g, b = colorsys.hls_to_rgb(h, clamped_l, s)
    return f"{round(r * 255):02X}{round(g * 255):02X}{round(b * 255):02X}"


__all__ = [
    "ThemeColors",
    "ThemeError",
    "apply_lum_mod",
    "apply_lum_off",
    "apply_theme_shade",
    "apply_theme_tint",
    "load_theme",
    "resolve_theme_color",
    "resolve_theme_font",
]
