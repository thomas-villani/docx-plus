"""Conditional table-style formatting — ``<w:tblStylePr>`` branches.

ECMA-376 17.7.6.5 lets a table style carry zero or more ``<w:tblStylePr>``
children, each with a ``w:type`` that names the conditional region
(``firstRow``, ``band1Horz``, ``nwCell``, etc.). The cascade resolver
applies these branches on top of the base table style in spec-defined
precedence order, picking the branches that match the target cell's
position within the table.

These tests verify the wiring end-to-end: build a document with a custom
table style, attach a table, then resolve formatting at different cell
positions and confirm the right branch was applied.
"""

from __future__ import annotations

import dataclasses
from typing import Any

import pytest
from docx import Document

from docx_plus.core.oxml import sub
from docx_plus.styles import TableContext, resolve_effective_formatting

# --------------------------------------------------------------------------
# Helpers — build a synthetic table style with conditional branches.
# --------------------------------------------------------------------------


def _add_table_style(
    doc: Document,
    style_id: str,
    *,
    base_rpr: dict[str, str] | None = None,
    branches: dict[str, dict[str, Any]] | None = None,
) -> None:
    """Append a ``w:style w:type="table"`` to the doc's styles part.

    Args:
        doc: Document to mutate.
        style_id: ``w:styleId`` for the new style.
        base_rpr: Optional dict of (rpr-child-tag -> attrs) for the
            base ``w:rPr`` of the style.
        branches: Mapping of conditional ``w:type`` → dict of
            ``w:rPr`` children, e.g. ``{"firstRow": {"w:b": None,
            "w:color": {"w:val": "FF0000"}}}``.
    """
    styles_el = doc.styles.element
    style_el = sub(styles_el, "w:style", **{"w:type": "table", "w:styleId": style_id})
    sub(style_el, "w:name", **{"w:val": style_id})

    if base_rpr:
        base = sub(style_el, "w:rPr")
        for tag, attrs in base_rpr.items():
            sub(base, tag, **(attrs or {}))

    if branches:
        for cond_type, rpr_children in branches.items():
            branch = sub(style_el, "w:tblStylePr", **{"w:type": cond_type})
            rpr = sub(branch, "w:rPr")
            for tag, attrs in rpr_children.items():
                sub(rpr, tag, **(attrs or {}))


def _add_table_with_style(doc: Document, style_id: str, rows: int, cols: int):
    """Add a fresh table to ``doc`` with ``style_id`` applied to ``tblPr``."""
    table = doc.add_table(rows=rows, cols=cols)
    tbl_pr = table._tbl.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblPr")
    if tbl_pr is None:
        tbl_pr = sub(table._tbl, "w:tblPr")
    sub(tbl_pr, "w:tblStyle", **{"w:val": style_id})
    return table


# --------------------------------------------------------------------------
# TableContext — defaults and field values.
# --------------------------------------------------------------------------


def test_table_context_defaults_are_all_false() -> None:
    ctx = TableContext()
    assert ctx.is_first_row is False
    assert ctx.is_last_row is False
    assert ctx.is_first_col is False
    assert ctx.is_last_col is False
    assert ctx.is_band_row is False
    assert ctx.is_band_col is False
    assert ctx.is_band2_row is False
    assert ctx.is_band2_col is False


def test_table_context_is_frozen() -> None:
    """TableContext is immutable (frozen dataclass)."""
    ctx = TableContext(is_first_row=True)
    with pytest.raises(dataclasses.FrozenInstanceError):
        ctx.is_first_row = False  # type: ignore[misc]


# --------------------------------------------------------------------------
# Auto-derive — TableContext is built from the cell's table position.
# --------------------------------------------------------------------------


def test_first_row_branch_applies_to_top_cell() -> None:
    """A 3x3 table with a ``firstRow`` branch sets bold on row 0 cells."""
    doc = Document()
    _add_table_style(
        doc,
        "ConditionalA",
        branches={"firstRow": {"w:b": None}},
    )
    table = _add_table_with_style(doc, "ConditionalA", rows=3, cols=3)

    top_cell = table.rows[0].cells[0]
    middle_cell = table.rows[1].cells[1]

    assert resolve_effective_formatting(top_cell).bold is True
    assert resolve_effective_formatting(middle_cell).bold is None


def test_last_row_branch_applies_to_bottom_cell() -> None:
    doc = Document()
    _add_table_style(
        doc,
        "ConditionalB",
        branches={"lastRow": {"w:i": None}},
    )
    table = _add_table_with_style(doc, "ConditionalB", rows=3, cols=3)

    last_cell = table.rows[2].cells[1]
    middle_cell = table.rows[1].cells[1]

    assert resolve_effective_formatting(last_cell).italic is True
    assert resolve_effective_formatting(middle_cell).italic is None


def test_band1_horz_applies_to_odd_rows() -> None:
    """``band1Horz`` colors rows 1, 3, ... (0-indexed)."""
    doc = Document()
    _add_table_style(
        doc,
        "Zebra",
        branches={"band1Horz": {"w:color": {"w:val": "FF0000"}}},
    )
    table = _add_table_with_style(doc, "Zebra", rows=4, cols=2)

    row_0_cell = table.rows[0].cells[0]
    row_1_cell = table.rows[1].cells[0]

    assert resolve_effective_formatting(row_0_cell).color_rgb is None
    assert resolve_effective_formatting(row_1_cell).color_rgb == "FF0000"


def test_band2_horz_applies_to_even_data_rows() -> None:
    """``band2Horz`` colors rows 2, 4, ... (complement of band1) — H5 regression.

    With the default band-size of 1, row 0 is the firstRow zone (skipped),
    rows 1/3/5 are band1, rows 2/4 are band2.
    """
    doc = Document()
    _add_table_style(
        doc,
        "ZebraBand2",
        branches={"band2Horz": {"w:color": {"w:val": "00FFFF"}}},
    )
    table = _add_table_with_style(doc, "ZebraBand2", rows=5, cols=2)

    row_0_cell = table.rows[0].cells[0]  # firstRow zone, not band
    row_1_cell = table.rows[1].cells[0]  # band1
    row_2_cell = table.rows[2].cells[0]  # band2
    row_3_cell = table.rows[3].cells[0]  # band1
    row_4_cell = table.rows[4].cells[0]  # band2

    assert resolve_effective_formatting(row_0_cell).color_rgb is None
    assert resolve_effective_formatting(row_1_cell).color_rgb is None
    assert resolve_effective_formatting(row_2_cell).color_rgb == "00FFFF"
    assert resolve_effective_formatting(row_3_cell).color_rgb is None
    assert resolve_effective_formatting(row_4_cell).color_rgb == "00FFFF"


def test_band2_vert_applies_to_even_data_columns() -> None:
    doc = Document()
    _add_table_style(
        doc,
        "VerticalBand2",
        branches={"band2Vert": {"w:color": {"w:val": "AA00AA"}}},
    )
    table = _add_table_with_style(doc, "VerticalBand2", rows=2, cols=5)

    col_0 = table.rows[0].cells[0]  # firstCol zone
    col_1 = table.rows[0].cells[1]  # band1
    col_2 = table.rows[0].cells[2]  # band2
    col_3 = table.rows[0].cells[3]  # band1
    col_4 = table.rows[0].cells[4]  # band2

    assert resolve_effective_formatting(col_0).color_rgb is None
    assert resolve_effective_formatting(col_1).color_rgb is None
    assert resolve_effective_formatting(col_2).color_rgb == "AA00AA"
    assert resolve_effective_formatting(col_3).color_rgb is None
    assert resolve_effective_formatting(col_4).color_rgb == "AA00AA"


def test_tbl_style_row_band_size_two_groups_rows_in_pairs() -> None:
    """``tblStyleRowBandSize="2"`` makes each band span two rows — H4 regression.

    Layout (band-size 2, no firstRow override): row 0 skipped, rows 1-2 →
    band1, rows 3-4 → band2, rows 5-6 → band1, ...
    """
    doc = Document()
    _add_table_style(
        doc,
        "Pairs",
        branches={
            "band1Horz": {"w:color": {"w:val": "111111"}},
            "band2Horz": {"w:color": {"w:val": "222222"}},
        },
    )
    table = _add_table_with_style(doc, "Pairs", rows=7, cols=1)
    # Set the band size on the table instance's own tblPr.
    tbl_pr = table._tbl.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblPr")
    assert tbl_pr is not None
    sub(tbl_pr, "w:tblStyleRowBandSize", **{"w:val": "2"})

    expected = [
        None,  # row 0 (firstRow zone)
        "111111",  # row 1 — band1 stripe 0
        "111111",  # row 2 — band1 stripe 0
        "222222",  # row 3 — band2 stripe 1
        "222222",  # row 4 — band2 stripe 1
        "111111",  # row 5 — band1 stripe 2
        "111111",  # row 6 — band1 stripe 2
    ]
    for row_idx, want in enumerate(expected):
        got = resolve_effective_formatting(table.rows[row_idx].cells[0]).color_rgb
        assert got == want, f"row {row_idx}: got {got!r}, want {want!r}"


def test_tbl_style_col_band_size_three_groups_columns_in_triples() -> None:
    doc = Document()
    _add_table_style(
        doc,
        "Triples",
        branches={"band1Vert": {"w:color": {"w:val": "333333"}}},
    )
    table = _add_table_with_style(doc, "Triples", rows=1, cols=7)
    tbl_pr = table._tbl.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblPr")
    assert tbl_pr is not None
    sub(tbl_pr, "w:tblStyleColBandSize", **{"w:val": "3"})

    # firstCol skipped; cols 1-3 → band1 stripe 0; cols 4-6 → band2 stripe 1.
    expected = [None, "333333", "333333", "333333", None, None, None]
    for col_idx, want in enumerate(expected):
        got = resolve_effective_formatting(table.rows[0].cells[col_idx]).color_rgb
        assert got == want, f"col {col_idx}: got {got!r}, want {want!r}"


def test_band1_vert_applies_to_odd_columns() -> None:
    doc = Document()
    _add_table_style(
        doc,
        "Vertical",
        branches={"band1Vert": {"w:color": {"w:val": "00FF00"}}},
    )
    table = _add_table_with_style(doc, "Vertical", rows=2, cols=4)

    col_0_cell = table.rows[0].cells[0]
    col_1_cell = table.rows[0].cells[1]

    assert resolve_effective_formatting(col_0_cell).color_rgb is None
    assert resolve_effective_formatting(col_1_cell).color_rgb == "00FF00"


# --------------------------------------------------------------------------
# Precedence — more-specific types override less-specific ones.
# --------------------------------------------------------------------------


def test_first_row_overrides_band1_horz() -> None:
    """``firstRow`` is later in spec order than ``band1Horz``; it wins.

    Even though row 1 (0-indexed) is a band row, the test checks row 0:
    row 0 matches ``firstRow`` only, not ``band1Horz``. Verify the
    first-row branch fires.
    """
    doc = Document()
    _add_table_style(
        doc,
        "Mixed",
        branches={
            "band1Horz": {"w:color": {"w:val": "111111"}},
            "firstRow": {"w:color": {"w:val": "222222"}},
        },
    )
    table = _add_table_with_style(doc, "Mixed", rows=3, cols=2)

    top_cell = table.rows[0].cells[0]
    resolved = resolve_effective_formatting(top_cell)
    assert resolved.color_rgb == "222222"


def test_first_col_overrides_first_row_at_corner_without_corner_branch() -> None:
    """Column branches win at row/col intersections — regression for H1.

    Per ECMA-376 17.7.6.5, the application order is rows → cols → corners,
    so a cell that matches both ``firstRow`` and ``firstCol`` (with no
    ``nwCell`` branch defined) must resolve to ``firstCol``'s properties.
    """
    doc = Document()
    _add_table_style(
        doc,
        "RowVsCol",
        branches={
            "firstRow": {"w:color": {"w:val": "AAAAAA"}},
            "firstCol": {"w:color": {"w:val": "BBBBBB"}},
        },
    )
    table = _add_table_with_style(doc, "RowVsCol", rows=3, cols=3)

    nw_cell = table.rows[0].cells[0]  # matches both firstRow and firstCol
    ne_cell = table.rows[0].cells[2]  # matches firstRow only
    sw_cell = table.rows[2].cells[0]  # matches firstCol only

    assert resolve_effective_formatting(nw_cell).color_rgb == "BBBBBB"  # firstCol wins
    assert resolve_effective_formatting(ne_cell).color_rgb == "AAAAAA"  # firstRow only
    assert resolve_effective_formatting(sw_cell).color_rgb == "BBBBBB"  # firstCol only


def test_single_row_table_last_row_overrides_first_row() -> None:
    """A 1-row table matches both firstRow and lastRow — lastRow wins."""
    doc = Document()
    _add_table_style(
        doc,
        "OneRow",
        branches={
            "firstRow": {"w:color": {"w:val": "111111"}},
            "lastRow": {"w:color": {"w:val": "222222"}},
        },
    )
    table = _add_table_with_style(doc, "OneRow", rows=1, cols=3)

    only_cell = table.rows[0].cells[1]
    assert resolve_effective_formatting(only_cell).color_rgb == "222222"


def test_single_column_table_last_col_overrides_first_col() -> None:
    """A 1-column table matches both firstCol and lastCol — lastCol wins."""
    doc = Document()
    _add_table_style(
        doc,
        "OneCol",
        branches={
            "firstCol": {"w:color": {"w:val": "333333"}},
            "lastCol": {"w:color": {"w:val": "444444"}},
        },
    )
    table = _add_table_with_style(doc, "OneCol", rows=3, cols=1)

    only_cell = table.rows[1].cells[0]
    assert resolve_effective_formatting(only_cell).color_rgb == "444444"


def test_corner_overrides_first_row() -> None:
    """``nwCell`` (corner) is later in spec order than ``firstRow``.

    The top-left cell matches both ``firstRow`` and ``nwCell``; the
    corner branch must override.
    """
    doc = Document()
    _add_table_style(
        doc,
        "Corners",
        branches={
            "firstRow": {"w:color": {"w:val": "AAAAAA"}},
            "nwCell": {"w:color": {"w:val": "BBBBBB"}},
        },
    )
    table = _add_table_with_style(doc, "Corners", rows=3, cols=3)

    nw_cell = table.rows[0].cells[0]
    ne_cell = table.rows[0].cells[2]

    assert resolve_effective_formatting(nw_cell).color_rgb == "BBBBBB"
    # NE matches firstRow but not nwCell — falls back to firstRow color.
    assert resolve_effective_formatting(ne_cell).color_rgb == "AAAAAA"


def test_child_base_overrides_parent_conditional_branch() -> None:
    """H9 regression: per-level base + conditional must interleave.

    Per ECMA-376 17.7.6.5 each style level computes (base then matching
    conditionals); the resulting per-level state then cascades child-
    over-parent. So a child style's BASE rPr must override a parent's
    matching conditional branch. The buggy implementation walked the
    whole chain for base first, then the whole chain for conditionals,
    inverting this at the parent/child boundary.
    """
    doc = Document()
    # Parent: firstRow branch sets color to ORANGE.
    _add_table_style(
        doc,
        "ParentWithFirstRow",
        branches={"firstRow": {"w:color": {"w:val": "FFA500"}}},
    )
    # Child basedOn parent: base rPr sets color to GREEN. No own firstRow.
    styles_el = doc.styles.element
    child = sub(
        styles_el,
        "w:style",
        **{"w:type": "table", "w:styleId": "ChildBaseGreen"},
    )
    sub(child, "w:name", **{"w:val": "ChildBaseGreen"})
    sub(child, "w:basedOn", **{"w:val": "ParentWithFirstRow"})
    cs_rpr = sub(child, "w:rPr")
    sub(cs_rpr, "w:color", **{"w:val": "00FF00"})

    table = _add_table_with_style(doc, "ChildBaseGreen", rows=3, cols=2)
    top_cell = table.rows[0].cells[0]
    resolved = resolve_effective_formatting(top_cell)
    # Order at row 0: parent base (none) → parent firstRow (ORANGE) →
    # child base (GREEN) → child firstRow (none). GREEN wins.
    assert resolved.color_rgb == "00FF00"


def test_whole_table_underlies_all_branches() -> None:
    """``wholeTable`` always applies; more-specific branches sit on top."""
    doc = Document()
    _add_table_style(
        doc,
        "BaseAndFirst",
        branches={
            "wholeTable": {"w:sz": {"w:val": "20"}},  # 10pt
            "firstRow": {"w:b": None},
        },
    )
    table = _add_table_with_style(doc, "BaseAndFirst", rows=2, cols=2)

    top_cell = table.rows[0].cells[0]
    middle_cell = table.rows[1].cells[1]

    # Top cell has firstRow's bold and wholeTable's size.
    top = resolve_effective_formatting(top_cell)
    assert top.bold is True
    assert top.font_size == 10.0
    # Middle cell only gets the wholeTable size.
    middle = resolve_effective_formatting(middle_cell)
    assert middle.bold is None
    assert middle.font_size == 10.0


# --------------------------------------------------------------------------
# Manual override — passing an explicit TableContext.
# --------------------------------------------------------------------------


def test_manual_table_context_overrides_auto_derived() -> None:
    """An explicit ``table_context`` arg supersedes auto-derivation."""
    doc = Document()
    _add_table_style(
        doc,
        "FirstRowBold",
        branches={"firstRow": {"w:b": None}},
    )
    table = _add_table_with_style(doc, "FirstRowBold", rows=3, cols=2)

    # Middle cell — auto-derived TableContext has is_first_row=False.
    middle_cell = table.rows[1].cells[0]
    auto = resolve_effective_formatting(middle_cell)
    assert auto.bold is None

    # Same cell, but the caller forces the first-row position.
    overridden = resolve_effective_formatting(
        middle_cell, table_context=TableContext(is_first_row=True)
    )
    assert overridden.bold is True


def test_non_table_target_has_no_conditional_formatting() -> None:
    """A regular paragraph never picks up table-style branches."""
    doc = Document()
    _add_table_style(
        doc,
        "ShouldNotApply",
        branches={"firstRow": {"w:b": None}},
    )
    p = doc.add_paragraph("standalone text")

    resolved = resolve_effective_formatting(p)
    assert resolved.bold is None


# --------------------------------------------------------------------------
# Paragraphs and runs inside table cells also receive conditional formatting.
# --------------------------------------------------------------------------


def test_conditional_formatting_reaches_paragraph_in_cell() -> None:
    doc = Document()
    _add_table_style(
        doc,
        "FirstRowItalic",
        branches={"firstRow": {"w:i": None}},
    )
    table = _add_table_with_style(doc, "FirstRowItalic", rows=2, cols=2)

    top_cell = table.rows[0].cells[0]
    top_cell.text = "header"
    para = top_cell.paragraphs[0]

    resolved = resolve_effective_formatting(para)
    assert resolved.italic is True


def test_conditional_formatting_reaches_run_in_cell() -> None:
    doc = Document()
    _add_table_style(
        doc,
        "FirstColBold",
        branches={"firstCol": {"w:b": None}},
    )
    table = _add_table_with_style(doc, "FirstColBold", rows=2, cols=3)

    left_cell = table.rows[1].cells[0]
    para = left_cell.paragraphs[0]
    run = para.add_run("data")

    resolved = resolve_effective_formatting(run)
    assert resolved.bold is True
