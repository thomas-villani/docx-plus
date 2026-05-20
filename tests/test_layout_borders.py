"""Tests for ``docx_plus.layout.set_page_borders`` and ``Border``."""

from __future__ import annotations

from pathlib import Path

from docx import Document

from docx_plus.core.ns import qn
from docx_plus.core.oxml import el, xpath
from docx_plus.layout import Border, set_page_borders


def _pg_borders(section):
    return section._sectPr.find(qn("w:pgBorders"))


# --------------------------------------------------------------------------
# Emission — element exists with one child per requested side.
# --------------------------------------------------------------------------


def test_set_page_borders_writes_element() -> None:
    doc = Document()
    set_page_borders(doc.sections[0], top=Border())
    assert _pg_borders(doc.sections[0]) is not None


def test_set_page_borders_all_four_sides() -> None:
    doc = Document()
    rule = Border(style="single", size=8, color="2F5496", space=24)
    set_page_borders(doc.sections[0], top=rule, bottom=rule, left=rule, right=rule)

    pg = _pg_borders(doc.sections[0])
    assert pg is not None
    side_tags = {child.tag for child in pg}
    assert side_tags == {qn(t) for t in ("w:top", "w:bottom", "w:left", "w:right")}

    for child in pg:
        assert child.get(qn("w:val")) == "single"
        assert child.get(qn("w:sz")) == "8"
        assert child.get(qn("w:color")) == "2F5496"
        assert child.get(qn("w:space")) == "24"


def test_set_page_borders_only_top() -> None:
    """Sides set to ``None`` are omitted from the emitted XML."""
    doc = Document()
    set_page_borders(doc.sections[0], top=Border(style="thick"))

    pg = _pg_borders(doc.sections[0])
    assert pg is not None
    side_tags = [child.tag for child in pg]
    assert side_tags == [qn("w:top")]
    assert pg.find(qn("w:top")).get(qn("w:val")) == "thick"


def test_border_defaults() -> None:
    """Bare ``Border()`` has the documented defaults."""
    b = Border()
    assert b.style == "single"
    assert b.size == 4
    assert b.color == "auto"
    assert b.space == 24


# --------------------------------------------------------------------------
# Replacement — second call overrides; all-None removes.
# --------------------------------------------------------------------------


def test_set_page_borders_replaces_existing() -> None:
    doc = Document()
    set_page_borders(doc.sections[0], top=Border(style="single"))
    set_page_borders(doc.sections[0], top=Border(style="double"))

    found = xpath(doc.sections[0]._sectPr, "./w:pgBorders")
    assert len(found) == 1
    assert found[0].find(qn("w:top")).get(qn("w:val")) == "double"


def test_set_page_borders_all_none_removes_element() -> None:
    doc = Document()
    set_page_borders(doc.sections[0], top=Border())
    assert _pg_borders(doc.sections[0]) is not None

    set_page_borders(doc.sections[0])  # all sides default to None
    assert _pg_borders(doc.sections[0]) is None


def test_set_page_borders_all_none_on_empty_is_noop() -> None:
    """All-``None`` on a section with no existing borders is a no-op."""
    doc = Document()
    set_page_borders(doc.sections[0])
    assert _pg_borders(doc.sections[0]) is None


# --------------------------------------------------------------------------
# Schema-strict insertion — pgBorders lands before its later siblings.
# --------------------------------------------------------------------------


def test_set_page_borders_lands_before_cols() -> None:
    doc = Document()
    sect_pr = doc.sections[0]._sectPr
    cols = el("w:cols", **{"w:num": "1"})
    sect_pr.append(cols)

    set_page_borders(doc.sections[0], top=Border())

    children = list(sect_pr)
    pg_idx = next(i for i, c in enumerate(children) if c.tag == qn("w:pgBorders"))
    cols_idx = next(i for i, c in enumerate(children) if c.tag == qn("w:cols"))
    assert pg_idx < cols_idx


def test_set_page_borders_lands_before_lnNumType() -> None:
    doc = Document()
    sect_pr = doc.sections[0]._sectPr
    ln = el("w:lnNumType", **{"w:countBy": "1"})
    sect_pr.append(ln)

    set_page_borders(doc.sections[0], top=Border())

    children = list(sect_pr)
    pg_idx = next(i for i, c in enumerate(children) if c.tag == qn("w:pgBorders"))
    ln_idx = next(i for i, c in enumerate(children) if c.tag == qn("w:lnNumType"))
    assert pg_idx < ln_idx


# --------------------------------------------------------------------------
# Round-trip — save / reopen preserves the element and attributes.
# --------------------------------------------------------------------------


def test_set_page_borders_round_trip(tmp_path: Path) -> None:
    doc = Document()
    set_page_borders(
        doc.sections[0],
        top=Border(style="double", size=12, color="FF0000", space=10),
        bottom=Border(style="single", size=4, color="auto", space=24),
    )
    out = tmp_path / "borders.docx"
    doc.save(str(out))

    reopened = Document(str(out))
    pg = _pg_borders(reopened.sections[0])
    assert pg is not None

    top = pg.find(qn("w:top"))
    assert top.get(qn("w:val")) == "double"
    assert top.get(qn("w:sz")) == "12"
    assert top.get(qn("w:color")) == "FF0000"
    assert top.get(qn("w:space")) == "10"

    bottom = pg.find(qn("w:bottom"))
    assert bottom.get(qn("w:val")) == "single"
    assert pg.find(qn("w:left")) is None
    assert pg.find(qn("w:right")) is None
