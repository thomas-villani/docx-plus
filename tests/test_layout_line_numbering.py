"""Tests for ``docx_plus.layout.set_line_numbering``."""

from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document

from docx_plus.core.ns import qn
from docx_plus.core.oxml import el, xpath
from docx_plus.layout import set_line_numbering


def _ln(section):
    return section._sectPr.find(qn("w:lnNumType"))


# --------------------------------------------------------------------------
# Emission — element shape and attribute values.
# --------------------------------------------------------------------------


def test_set_line_numbering_writes_element() -> None:
    doc = Document()
    set_line_numbering(doc.sections[0])
    assert _ln(doc.sections[0]) is not None


def test_set_line_numbering_default_attrs() -> None:
    doc = Document()
    set_line_numbering(doc.sections[0])
    node = _ln(doc.sections[0])
    assert node.get(qn("w:countBy")) == "1"
    assert node.get(qn("w:start")) == "1"
    assert node.get(qn("w:restart")) == "newPage"
    # distance is intentionally omitted so Word picks its built-in default
    assert node.get(qn("w:distance")) is None


def test_set_line_numbering_custom_values() -> None:
    doc = Document()
    set_line_numbering(
        doc.sections[0],
        count_by=5,
        restart="continuous",
        start=11,
        distance=360,
    )
    node = _ln(doc.sections[0])
    assert node.get(qn("w:countBy")) == "5"
    assert node.get(qn("w:restart")) == "continuous"
    assert node.get(qn("w:start")) == "11"
    assert node.get(qn("w:distance")) == "360"


# --------------------------------------------------------------------------
# Replacement — second call overrides rather than stacks.
# --------------------------------------------------------------------------


def test_set_line_numbering_is_idempotent() -> None:
    doc = Document()
    set_line_numbering(doc.sections[0], count_by=1)
    set_line_numbering(doc.sections[0], count_by=10)
    found = xpath(doc.sections[0]._sectPr, "./w:lnNumType")
    assert len(found) == 1
    assert found[0].get(qn("w:countBy")) == "10"


# --------------------------------------------------------------------------
# Validation — argument guards.
# --------------------------------------------------------------------------


def test_set_line_numbering_rejects_zero_count_by() -> None:
    doc = Document()
    with pytest.raises(ValueError, match="count_by >= 1"):
        set_line_numbering(doc.sections[0], count_by=0)


def test_set_line_numbering_rejects_zero_start() -> None:
    doc = Document()
    with pytest.raises(ValueError, match="start >= 1"):
        set_line_numbering(doc.sections[0], start=0)


def test_set_line_numbering_rejects_unknown_restart() -> None:
    doc = Document()
    with pytest.raises(ValueError, match="restart must be"):
        set_line_numbering(doc.sections[0], restart="bogus")  # type: ignore[arg-type]


# --------------------------------------------------------------------------
# Schema-strict insertion — lnNumType lands before its later siblings.
# --------------------------------------------------------------------------


def test_set_line_numbering_lands_before_cols() -> None:
    """``lnNumType`` is schema-positioned before ``cols``."""
    doc = Document()
    sect_pr = doc.sections[0]._sectPr
    # Pre-seed an empty `<w:cols>` so insert_before_first_anchor has a target.
    cols = el("w:cols", **{"w:num": "1"})
    sect_pr.append(cols)

    set_line_numbering(doc.sections[0], count_by=2)

    children = list(sect_pr)
    ln_idx = next(i for i, c in enumerate(children) if c.tag == qn("w:lnNumType"))
    cols_idx = next(i for i, c in enumerate(children) if c.tag == qn("w:cols"))
    assert ln_idx < cols_idx


# --------------------------------------------------------------------------
# Round-trip — save / reopen preserves the element.
# --------------------------------------------------------------------------


def test_set_line_numbering_round_trip(tmp_path: Path) -> None:
    doc = Document()
    set_line_numbering(doc.sections[0], count_by=5, restart="newSection", start=21)
    out = tmp_path / "ln.docx"
    doc.save(str(out))

    reopened = Document(str(out))
    node = _ln(reopened.sections[0])
    assert node is not None
    assert node.get(qn("w:countBy")) == "5"
    assert node.get(qn("w:restart")) == "newSection"
    assert node.get(qn("w:start")) == "21"
