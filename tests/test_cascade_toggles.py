"""Toggle property XOR semantics — the highest-risk part of the cascade.

ECMA-376 17.7.3 specifies that certain run properties (bold, italic, caps,
smallCaps, strike, vanish, ...) XOR through the basedOn chain rather than
override. An explicit ``w:val="false"`` resets parity to false from that
layer onward. These cases will look pedantic; they are exactly the bugs that
silently survive otherwise. IMPLEMENTATION.md §5 lists them verbatim.
"""

from __future__ import annotations

from typing import Any

from docx import Document

from docx_plus.core.oxml import sub
from docx_plus.styles.inspect import resolve_effective_formatting


def _add_paragraph_style(
    doc: Document,
    style_id: str,
    *,
    based_on: str | None = None,
    rpr_children: list[tuple[str, dict[str, str] | None]] | None = None,
) -> None:
    """Append a paragraph style to the doc's styles part.

    Args:
        doc: Document to mutate.
        style_id: ``w:styleId`` for the new style.
        based_on: Optional basedOn target.
        rpr_children: List of ``(tag, attrs|None)`` to add inside ``w:rPr``.
    """
    styles_el = doc.styles.element
    s = sub(styles_el, "w:style", **{"w:type": "paragraph", "w:styleId": style_id})
    sub(s, "w:name", **{"w:val": style_id})
    if based_on is not None:
        sub(s, "w:basedOn", **{"w:val": based_on})
    if rpr_children:
        rpr = sub(s, "w:rPr")
        for tag, attrs in rpr_children:
            kwargs: dict[str, Any] = attrs or {}
            sub(rpr, tag, **kwargs)


def _styled_paragraph(doc: Document, style_id: str) -> Any:
    p = doc.add_paragraph()
    ppr = p._p.get_or_add_pPr()
    sub(ppr, "w:pStyle", **{"w:val": style_id})
    return p


# --------------------------------------------------------------------------
# Case 1: Style defines bold, no further override -> bold.
# --------------------------------------------------------------------------


def test_single_style_bold_is_bold() -> None:
    doc = Document()
    _add_paragraph_style(doc, "BoldOnly", rpr_children=[("w:b", None)])
    p = _styled_paragraph(doc, "BoldOnly")

    resolved = resolve_effective_formatting(p)
    assert resolved.bold is True


# --------------------------------------------------------------------------
# Case 2: Style A bold + B basedOn A also bold -> NOT bold (XOR).
# --------------------------------------------------------------------------


def test_bold_xor_through_basedon_chain() -> None:
    doc = Document()
    _add_paragraph_style(doc, "ABold", rpr_children=[("w:b", None)])
    _add_paragraph_style(
        doc, "BBoldChild", based_on="ABold", rpr_children=[("w:b", None)]
    )
    p = _styled_paragraph(doc, "BBoldChild")

    resolved = resolve_effective_formatting(p)
    assert resolved.bold is False


# --------------------------------------------------------------------------
# Case 3: Style A bold + B basedOn A with `w:val="false"` -> false (reset).
# --------------------------------------------------------------------------


def test_explicit_false_resets_toggle() -> None:
    doc = Document()
    _add_paragraph_style(doc, "ABold2", rpr_children=[("w:b", None)])
    _add_paragraph_style(
        doc,
        "BUnboldChild",
        based_on="ABold2",
        rpr_children=[("w:b", {"w:val": "false"})],
    )
    p = _styled_paragraph(doc, "BUnboldChild")

    resolved = resolve_effective_formatting(p)
    assert resolved.bold is False


# --------------------------------------------------------------------------
# Case 4: Direct bold on a non-bold style -> bold (XOR from None -> True).
# --------------------------------------------------------------------------


def test_direct_bold_on_unbold_style() -> None:
    doc = Document()
    _add_paragraph_style(doc, "PlainStyle")  # no w:b
    p = _styled_paragraph(doc, "PlainStyle")
    r = p.add_run("text")
    r_rpr = r._r.get_or_add_rPr()
    sub(r_rpr, "w:b")

    resolved = resolve_effective_formatting(r)
    assert resolved.bold is True


# --------------------------------------------------------------------------
# Case 5: Direct `w:val="false"` on a bold style -> not bold.
# --------------------------------------------------------------------------


def test_direct_unbold_on_bold_style() -> None:
    doc = Document()
    _add_paragraph_style(doc, "BoldStyleA", rpr_children=[("w:b", None)])
    p = _styled_paragraph(doc, "BoldStyleA")
    r = p.add_run("text")
    r_rpr = r._r.get_or_add_rPr()
    sub(r_rpr, "w:b", **{"w:val": "false"})

    resolved = resolve_effective_formatting(r)
    assert resolved.bold is False


# --------------------------------------------------------------------------
# Additional verification: three-level chain parity.
# --------------------------------------------------------------------------


def test_three_level_xor_parity() -> None:
    """Bold at three levels: True XOR True XOR True -> True."""
    doc = Document()
    _add_paragraph_style(doc, "L1", rpr_children=[("w:b", None)])
    _add_paragraph_style(doc, "L2", based_on="L1", rpr_children=[("w:b", None)])
    _add_paragraph_style(doc, "L3", based_on="L2", rpr_children=[("w:b", None)])
    p = _styled_paragraph(doc, "L3")

    resolved = resolve_effective_formatting(p)
    assert resolved.bold is True


def test_italic_xor_alongside_bold_independent() -> None:
    """Each toggle property tracks parity independently."""
    doc = Document()
    _add_paragraph_style(
        doc, "BoldOnce", rpr_children=[("w:b", None), ("w:i", None)]
    )
    _add_paragraph_style(
        doc, "BoldTwice", based_on="BoldOnce", rpr_children=[("w:b", None)]
    )
    p = _styled_paragraph(doc, "BoldTwice")

    resolved = resolve_effective_formatting(p)
    assert resolved.bold is False  # XOR twice
    assert resolved.italic is True  # Set once


def test_value_zero_treated_as_explicit_false() -> None:
    """`w:val="0"` is the legacy form of `w:val="false"` per the schema."""
    doc = Document()
    _add_paragraph_style(doc, "BoldZ", rpr_children=[("w:b", None)])
    _add_paragraph_style(
        doc,
        "UnboldZ",
        based_on="BoldZ",
        rpr_children=[("w:b", {"w:val": "0"})],
    )
    p = _styled_paragraph(doc, "UnboldZ")

    resolved = resolve_effective_formatting(p)
    assert resolved.bold is False


def test_value_true_explicit_xor_flips() -> None:
    """`w:val="true"` is equivalent to no val for XOR purposes."""
    doc = Document()
    _add_paragraph_style(
        doc, "BoldTrue", rpr_children=[("w:b", {"w:val": "true"})]
    )
    p = _styled_paragraph(doc, "BoldTrue")

    resolved = resolve_effective_formatting(p)
    assert resolved.bold is True
