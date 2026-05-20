"""Tests for ``docx_plus.publishing.add_toc``."""

from __future__ import annotations

from pathlib import Path

from docx import Document

from docx_plus.core.ns import qn
from docx_plus.core.oxml import xpath
from docx_plus.publishing import add_toc


def _instruction(field_run):
    """Extract the instruction string from the field that ``field_run`` opens."""
    p = field_run.getparent()
    instr = p.find(qn("w:r") + "/" + qn("w:instrText"))
    # Fallback: search any run for instrText (the begin-run has no instrText).
    if instr is None:
        for r in p.findall(qn("w:r")):
            t = r.find(qn("w:instrText"))
            if t is not None:
                return t.text
        return None
    return instr.text


# --------------------------------------------------------------------------
# Emission — instruction string assembly.
# --------------------------------------------------------------------------


def test_add_toc_defaults_match_word_output() -> None:
    doc = Document()
    field = add_toc(doc.add_paragraph())
    instr = _instruction(field)
    assert instr is not None
    assert 'TOC \\o "1-3"' in instr
    assert "\\h" in instr
    assert "\\z" in instr
    assert "\\u" in instr
    assert "\\n" not in instr  # page_numbers default True


def test_add_toc_custom_levels() -> None:
    doc = Document()
    field = add_toc(doc.add_paragraph(), levels=(2, 5))
    instr = _instruction(field)
    assert '\\o "2-5"' in instr


def test_add_toc_no_hyperlink() -> None:
    doc = Document()
    field = add_toc(doc.add_paragraph(), hyperlink=False)
    instr = _instruction(field)
    assert "\\h" not in instr


def test_add_toc_no_page_numbers_emits_n_switch() -> None:
    doc = Document()
    field = add_toc(doc.add_paragraph(), page_numbers=False)
    instr = _instruction(field)
    assert "\\n" in instr


# --------------------------------------------------------------------------
# Field structure — five-run complex-field sequence.
# --------------------------------------------------------------------------


def test_add_toc_writes_full_complex_field() -> None:
    doc = Document()
    p = doc.add_paragraph()
    add_toc(p)

    fld_chars = xpath(p._p, ".//w:fldChar")
    char_types = [c.get(qn("w:fldCharType")) for c in fld_chars]
    assert char_types == ["begin", "separate", "end"]


def test_add_toc_returns_begin_run() -> None:
    doc = Document()
    field = add_toc(doc.add_paragraph())
    assert field.tag == qn("w:r")
    fld_char = field.find(qn("w:fldChar"))
    assert fld_char is not None
    assert fld_char.get(qn("w:fldCharType")) == "begin"


# --------------------------------------------------------------------------
# Round-trip — save / reopen preserves the field.
# --------------------------------------------------------------------------


def test_add_toc_round_trip(tmp_path: Path) -> None:
    doc = Document()
    add_toc(doc.add_paragraph(), levels=(1, 4), page_numbers=False)
    out = tmp_path / "toc.docx"
    doc.save(str(out))

    reopened = Document(str(out))
    body = reopened.element.body
    instr_runs = xpath(body, ".//w:instrText")
    assert len(instr_runs) == 1
    text = instr_runs[0].text
    assert 'TOC \\o "1-4"' in text
    assert "\\n" in text
