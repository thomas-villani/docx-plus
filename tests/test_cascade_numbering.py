"""Tests for cascade layer 4 (numbering) in :func:`resolve_effective_formatting`.

The ``numbered`` fixture has one paragraph whose ``w:numPr`` references a
custom ``abstractNum`` carrying both pPr (indent 720 left, -360 first-line)
and rPr (bold) at ``lvl[0]``. These tests verify the resolver picks all of
that up, attributes it to the ``numbering`` layer, and degrades gracefully
when the references can't be followed.
"""

from __future__ import annotations

from pathlib import Path

from docx import Document

from docx_plus.core.ns import qn
from docx_plus.styles import resolve_effective_formatting


def test_numbering_layer_resolves_indent_and_bold(numbered_docx_path: Path) -> None:
    """Both pPr (indent) and rPr (bold) on lvl[0] flow into ResolvedFormatting."""
    doc = Document(str(numbered_docx_path))
    resolved = resolve_effective_formatting(doc.paragraphs[0])
    assert resolved.num_id == 100
    assert resolved.num_level == 0
    assert resolved.indent_left == 720
    assert resolved.indent_first_line == -360
    assert resolved.bold is True


def test_numbering_provenance_marks_layer(numbered_docx_path: Path) -> None:
    """Provenance attributes the indent + bold to the ``numbering`` layer."""
    doc = Document(str(numbered_docx_path))
    resolved = resolve_effective_formatting(doc.paragraphs[0], include_provenance=True)
    prov = resolved.provenance or {}
    assert prov["indent_left"].layer == "numbering"
    assert prov["bold"].layer == "numbering"
    assert prov["num_id"].layer == "numbering"


def test_numbering_unknown_num_id_skips_silently(numbered_docx_path: Path) -> None:
    """A numPr that references an absent num is non-fatal — the layer just no-ops."""
    doc = Document(str(numbered_docx_path))
    para = doc.paragraphs[0]
    num_pr = para._p.find(f"./{qn('w:pPr')}/{qn('w:numPr')}")
    assert num_pr is not None
    num_id_el = num_pr.find(qn("w:numId"))
    assert num_id_el is not None
    num_id_el.set(qn("w:val"), "9999")  # not in numbering.xml
    resolved = resolve_effective_formatting(para)
    # num_id is still recorded from the paragraph's numPr, but no pPr/rPr
    # flows in because the resolver can't find the abstractNum.
    assert resolved.num_id == 9999
    assert resolved.bold is None


def test_numbering_unparsable_numid_is_ignored(numbered_docx_path: Path) -> None:
    """A non-numeric w:numId/@w:val short-circuits cleanly without raising."""
    doc = Document(str(numbered_docx_path))
    para = doc.paragraphs[0]
    num_pr = para._p.find(f"./{qn('w:pPr')}/{qn('w:numPr')}")
    assert num_pr is not None
    num_id_el = num_pr.find(qn("w:numId"))
    assert num_id_el is not None
    num_id_el.set(qn("w:val"), "not-a-number")
    resolved = resolve_effective_formatting(para)
    assert resolved.num_id is None
    assert resolved.bold is None


def test_numbering_unparsable_ilvl_falls_back_to_zero(
    numbered_docx_path: Path,
) -> None:
    """A bogus ilvl falls back to level 0 rather than raising."""
    doc = Document(str(numbered_docx_path))
    para = doc.paragraphs[0]
    num_pr = para._p.find(f"./{qn('w:pPr')}/{qn('w:numPr')}")
    assert num_pr is not None
    ilvl_el = num_pr.find(qn("w:ilvl"))
    assert ilvl_el is not None
    ilvl_el.set(qn("w:val"), "garbage")
    resolved = resolve_effective_formatting(para)
    # Level fallback kicked in; the indent + bold at lvl[0] still apply.
    assert resolved.num_level == 0
    assert resolved.bold is True


def test_numbering_missing_numid_attribute_short_circuits(
    numbered_docx_path: Path,
) -> None:
    """w:numId without a w:val attribute exits early without setting num_id."""
    doc = Document(str(numbered_docx_path))
    para = doc.paragraphs[0]
    num_pr = para._p.find(f"./{qn('w:pPr')}/{qn('w:numPr')}")
    assert num_pr is not None
    num_id_el = num_pr.find(qn("w:numId"))
    assert num_id_el is not None
    # Remove the val attribute entirely.
    del num_id_el.attrib[qn("w:val")]
    resolved = resolve_effective_formatting(para)
    assert resolved.num_id is None


def test_numbering_paragraph_without_numpr_skips_layer(
    multistyle_docx_path: Path,
) -> None:
    """A plain paragraph (no w:numPr) doesn't trigger the numbering layer."""
    doc = Document(str(multistyle_docx_path))
    resolved = resolve_effective_formatting(doc.paragraphs[0], include_provenance=True)
    assert resolved.num_id is None
    prov = resolved.provenance or {}
    # No numbering-layer entries appear.
    assert not any(src.layer == "numbering" for src in prov.values())
