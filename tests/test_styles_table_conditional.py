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

from typing import Any

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


def test_table_context_is_frozen() -> None:
    """TableContext is immutable (frozen dataclass)."""
    ctx = TableContext(is_first_row=True)
    try:
        ctx.is_first_row = False  # type: ignore[misc]
    except Exception:
        return
    raise AssertionError("expected TableContext to reject attribute assignment")


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
