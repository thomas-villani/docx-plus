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
They are correspondingly **not** verified against Word: no cascade input
can produce one, so there is nothing to render and compare.

Which scheme slot a name resolves to is per-document. ``settings.xml``
carries a ``<w:clrSchemeMapping>`` that redirects the *semantic* names —
``text1``, ``background1``, ``accent1``, ``hyperlink``, … — onto scheme
slots; a dark-themed template swaps ``t1`` and ``bg1`` so ``text1`` renders
white. The direct slot names (``dark1`` / ``light1`` / ``dark2`` /
``light2``) are never redirected, so the mapping is not a rename of the
scheme. Measured against Word.

Arithmetic here is exact (:class:`fractions.Fraction`), not floating
point. These transforms land on integer boundaries often enough that it
matters: ``1 - 0xE6/255`` is 0.09803921568627449 in binary floating point,
and 255 times that is 24.999999999999996, one below the 25 Word paints.
Exactness also makes the RGB -> HSL -> RGB round-trip lossless, which is
what lets the final channel be truncated (matching Word) without a no-op
transform changing the colour.

The module is read-only; writing themes is a v0.2 non-goal (SPEC §1).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from fractions import Fraction
from typing import TYPE_CHECKING

from lxml import etree

from docx_plus.core import DocxPlusError
from docx_plus.core.ns import A, qn
from docx_plus.core.oxml import xpath

if TYPE_CHECKING:
    from docx.document import Document


_HALF = Fraction(1, 2)

_THEME_RELTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"

# ECMA-376 ST_ColorSchemeIndex -> DrawingML clrScheme child element name.
# This is the *scheme slot* map and is fixed; which slot a given
# ``w:themeColor`` name lands in is a separate, per-document question — see
# _CLR_SCHEME_MAPPING_ATTR.
_SCHEME_INDEX_TO_KEY: dict[str, str] = {
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
}

# Word's ST_ThemeColor name -> the ``<w:clrSchemeMapping>`` attribute that
# says which scheme slot it resolves to.
#
# Only the *semantic* names are remapped. ``dark1`` / ``light1`` / ``dark2``
# / ``light2`` name scheme slots directly and are never redirected —
# measured against Word, which honours the mapping for ``text1`` and friends
# while leaving ``dark1`` alone in the same document.
_CLR_SCHEME_MAPPING_ATTR: dict[str, str] = {
    "text1": "t1",
    "background1": "bg1",
    "text2": "t2",
    "background2": "bg2",
    "accent1": "accent1",
    "accent2": "accent2",
    "accent3": "accent3",
    "accent4": "accent4",
    "accent5": "accent5",
    "accent6": "accent6",
    "hyperlink": "hyperlink",
    "followedHyperlink": "followedHyperlink",
}

# The mapping Word writes into a new document, and what it assumes for any
# attribute a ``<w:clrSchemeMapping>`` leaves out.
_DEFAULT_CLR_SCHEME_MAPPING: dict[str, str] = {
    "t1": "dark1",
    "bg1": "light1",
    "t2": "dark2",
    "bg2": "light2",
    "accent1": "accent1",
    "accent2": "accent2",
    "accent3": "accent3",
    "accent4": "accent4",
    "accent5": "accent5",
    "accent6": "accent6",
    "hyperlink": "hyperlink",
    "followedHyperlink": "followedHyperlink",
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
        mapping: The document's ``<w:clrSchemeMapping>`` — attribute name
            (``"t1"``, ``"bg1"``, ``"accent1"``, …) -> scheme slot. Defaults
            to what Word writes into a new document, so an absent or partial
            element behaves as Word treats it.
    """

    scheme: dict[str, str]
    fonts: dict[str, str] = field(default_factory=dict)
    mapping: dict[str, str] = field(default_factory=lambda: dict(_DEFAULT_CLR_SCHEME_MAPPING))

    def base(self, theme_name: str) -> str | None:
        """Return the unmodified hex color for a Word theme color name.

        Resolves through the document's ``<w:clrSchemeMapping>`` where one
        applies: ``text1`` means "whatever slot this document maps ``t1``
        to", which is ``dark1`` by default but need not be. The direct slot
        names (``dark1`` / ``light1`` / ``dark2`` / ``light2``) bypass the
        mapping.

        Args:
            theme_name: A value from ``ST_ThemeColor`` (e.g. ``"accent1"``,
                ``"text1"``).

        Returns:
            Uppercase ``RRGGBB`` hex string, or ``None`` if the name is not a
            recognized theme color or the underlying scheme entry is missing.
        """
        attr = _CLR_SCHEME_MAPPING_ATTR.get(theme_name)
        if attr is not None:
            slot = self.mapping.get(attr, _DEFAULT_CLR_SCHEME_MAPPING[attr])
        else:
            slot = theme_name
        key = _SCHEME_INDEX_TO_KEY.get(slot)
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
    return ThemeColors(
        scheme=_parse_clr_scheme(root),
        fonts=_parse_font_scheme(root),
        mapping=_read_clr_scheme_mapping(doc),
    )


def _read_clr_scheme_mapping(doc: Document) -> dict[str, str]:
    """Read ``<w:clrSchemeMapping>`` from ``settings.xml``.

    Word consults this to decide which scheme slot a ``w:themeColor`` name
    such as ``text1`` refers to; a dark-themed template swaps ``t1`` and
    ``bg1`` so that ``text1`` renders white. Attributes the element omits —
    and a document with no element or no settings part at all — fall back to
    Word's defaults.
    """
    mapping = dict(_DEFAULT_CLR_SCHEME_MAPPING)
    try:
        settings_el = doc.settings.element
    except Exception:  # pragma: no cover - a document with no settings part
        return mapping
    element = settings_el.find(qn("w:clrSchemeMapping"))
    if element is None:
        return mapping
    for attr in _DEFAULT_CLR_SCHEME_MAPPING:
        value = element.get(qn(f"w:{attr}"))
        if value in _SCHEME_INDEX_TO_KEY:
            mapping[attr] = value
    return mapping


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
    return _hls_to_hex(h, _lerp_to_white(lum, t), s)


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


def _lerp_to_white(lum: Fraction, tint: Fraction) -> Fraction:
    """``L * t + (1 - t)`` — kept exact so boundary values land where Word puts them."""
    return lum * tint + (Fraction(1) - tint)


def apply_lum_mod(hex_color: str, lum_mod: int) -> str:
    """Multiply L by ``lum_mod / 100000`` per ECMA-376 17.18.40.

    DrawingML transform values are percent thousandths: ``50000`` means 50%.

    Args:
        hex_color: Six-character hex color.
        lum_mod: Percent thousandths (e.g. ``50000`` for 50%).

    Returns:
        Uppercase ``RRGGBB`` hex string with L clamped to ``[0, 1]``.
    """
    factor = Fraction(lum_mod, 100000)
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
    delta = Fraction(lum_off, 100000)
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


def _parse_hex_byte(byte_str: str) -> Fraction:
    try:
        value = int(byte_str, 16)
    except (TypeError, ValueError) as exc:
        raise ThemeError(f"expected hex byte, got {byte_str!r}") from exc
    if not 0 <= value <= 0xFF:
        raise ThemeError(f"hex byte {byte_str!r} out of range")
    return Fraction(value, 255)


def _rgb_to_hls(hex_color: str) -> tuple[Fraction, Fraction, Fraction]:
    """Exact HSL, as Fractions in ``[0, 1]``, in colorsys's ``(h, l, s)`` order.

    Deliberately not :func:`colorsys.rgb_to_hls`: these transforms land on
    exact integer boundaries often enough that binary floating point changes
    the answer. ``1 - 0xE6/255`` is 0.09803921568627449, and 255 times that
    is 24.999999999999996 — which truncates to 24 where Word renders 25.
    """
    cleaned = hex_color.lstrip("#")
    if len(cleaned) != 6:
        raise ThemeError(f"expected 6-character hex color, got {hex_color!r}")
    try:
        r = int(cleaned[0:2], 16)
        g = int(cleaned[2:4], 16)
        b = int(cleaned[4:6], 16)
    except ValueError as exc:
        raise ThemeError(f"unparseable hex color {hex_color!r}") from exc

    rf, gf, bf = Fraction(r, 255), Fraction(g, 255), Fraction(b, 255)
    hi, lo = max(rf, gf, bf), min(rf, gf, bf)
    lum = (hi + lo) / 2
    if hi == lo:
        return Fraction(0), lum, Fraction(0)
    span = hi - lo
    sat = span / (hi + lo) if lum <= _HALF else span / (2 - hi - lo)
    if hi == rf:
        hue = (gf - bf) / span
    elif hi == gf:
        hue = 2 + (bf - rf) / span
    else:
        hue = 4 + (rf - gf) / span
    return (hue / 6) % 1, lum, sat


def _hue_to_channel(m1: Fraction, m2: Fraction, hue: Fraction) -> Fraction:
    hue = hue % 1
    if hue < Fraction(1, 6):
        return m1 + (m2 - m1) * 6 * hue
    if hue < _HALF:
        return m2
    if hue < Fraction(2, 3):
        return m1 + (m2 - m1) * (Fraction(2, 3) - hue) * 6
    return m1


def _hls_to_hex(h: Fraction, lum: Fraction, s: Fraction) -> str:
    """Convert back to hex, truncating each channel as Word does.

    Truncation is safe here only because the arithmetic is exact: an
    untransformed colour round-trips to channel values that are exact
    multiples of 1/255, so ``floor`` and ``round`` agree and a no-op
    transform stays a no-op. In floating point that invariant does not
    hold, which is why this module does not use :mod:`colorsys`.
    """
    lum = max(Fraction(0), min(Fraction(1), lum))
    if s == 0:
        channels = (lum, lum, lum)
    else:
        m2 = lum * (1 + s) if lum <= _HALF else lum + s - lum * s
        m1 = 2 * lum - m2
        channels = tuple(  # type: ignore[assignment]
            _hue_to_channel(m1, m2, h + offset)
            for offset in (Fraction(1, 3), Fraction(0), Fraction(-1, 3))
        )
    out = []
    for channel in channels:
        scaled = channel * 255
        out.append(max(0, min(255, scaled.numerator // scaled.denominator)))
    return "".join(f"{value:02X}" for value in out)


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
