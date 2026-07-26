"""Tests for ``docx_plus.core.borders`` — the shared ``CT_Border`` shape."""

from __future__ import annotations

import pytest

from docx_plus.core.borders import Border, border_attrs

# --------------------------------------------------------------------------
# Serialization — every border edge in the format takes the same four attrs.
# --------------------------------------------------------------------------


def test_border_attrs_maps_all_four_attributes() -> None:
    attrs = border_attrs(Border(style="double", size=12, color="2F5496", space=0))
    assert attrs == {
        "w:val": "double",
        "w:sz": "12",
        "w:color": "2F5496",
        "w:space": "0",
    }


def test_border_attrs_stringifies_numeric_fields() -> None:
    attrs = border_attrs(Border())
    assert attrs["w:sz"] == "4"
    assert attrs["w:space"] == "24"


# --------------------------------------------------------------------------
# Validation. The color check shipped in v0.2; size and style were promised
# by the docstring but unenforced until the promotion to core.
# --------------------------------------------------------------------------


@pytest.mark.parametrize("color", ["auto", "000000", "FFFFFF", "2f5496"])
def test_border_accepts_valid_colors(color: str) -> None:
    assert Border(color=color).color == color


@pytest.mark.parametrize("color", ["red", "#FF0000", "F00", "menu", ""])
def test_border_rejects_invalid_colors(color: str) -> None:
    with pytest.raises(ValueError, match="Border.color"):
        Border(color=color)


@pytest.mark.parametrize("size", [0, 4, 96])
def test_border_accepts_sizes_within_the_ecma_cap(size: int) -> None:
    assert Border(size=size).size == size


@pytest.mark.parametrize("size", [-1, 97, 1000])
def test_border_rejects_sizes_outside_the_ecma_cap(size: int) -> None:
    with pytest.raises(ValueError, match="eighths of a point"):
        Border(size=size)


@pytest.mark.parametrize("style", ["single", "none", "dotDotDash", "threeDEmboss"])
def test_border_accepts_st_border_names(style: str) -> None:
    assert Border(style=style).style == style


@pytest.mark.parametrize("style", ["1px solid", "#000", "", "double-line"])
def test_border_rejects_css_isms(style: str) -> None:
    # The full ST_Border enumeration runs past 200 clip-art values, so the
    # check is shape-based: catch the CSS habit without rejecting an exotic
    # but legitimate name.
    with pytest.raises(ValueError, match="Border.style"):
        Border(style=style)


def test_border_is_frozen() -> None:
    border = Border()
    with pytest.raises(AttributeError):
        border.size = 8  # type: ignore[misc]


def test_layout_still_re_exports_border() -> None:
    """``Border`` shipped from ``layout`` in v0.2; the name must keep working."""
    from docx_plus.layout import Border as LayoutBorder

    assert LayoutBorder is Border
