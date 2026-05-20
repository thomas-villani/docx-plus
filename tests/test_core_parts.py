"""Tests for ``docx_plus.core.parts`` — get-or-create plumbing for
separate OOXML parts (comments, footnotes, endnotes)."""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import PartFactory, XmlPart
from docx.parts.comments import CommentsPart

from docx_plus.core.ns import qn
from docx_plus.core.oxml import sub
from docx_plus.core.parts import (
    COMMENTS_SPEC,
    ENDNOTES_SPEC,
    FOOTNOTES_SPEC,
    _EndnotesPart,
    _FootnotesPart,
    get_or_create_part,
)

# --------------------------------------------------------------------------
# Part-class registration: footnote and endnote parts must round-trip with
# parsed XML rather than as raw blobs.
# --------------------------------------------------------------------------


def test_footnotes_part_class_registered() -> None:
    assert PartFactory.part_type_for[CT.WML_FOOTNOTES] is _FootnotesPart


def test_endnotes_part_class_registered() -> None:
    assert PartFactory.part_type_for[CT.WML_ENDNOTES] is _EndnotesPart


def test_comments_part_class_still_python_docx_default() -> None:
    """python-docx registers ``CommentsPart`` itself; we must not clobber it."""
    assert PartFactory.part_type_for[CT.WML_COMMENTS] is CommentsPart


# --------------------------------------------------------------------------
# Create paths: fresh document has none of the optional parts.
# --------------------------------------------------------------------------


def test_get_or_create_comments_creates_when_absent() -> None:
    doc = Document()
    part, root = get_or_create_part(doc, COMMENTS_SPEC)
    assert isinstance(part, CommentsPart)
    assert part.content_type == CT.WML_COMMENTS
    assert part.partname == "/word/comments.xml"
    assert root.tag == qn("w:comments")
    assert len(list(root)) == 0  # empty comments root


def test_get_or_create_footnotes_creates_when_absent() -> None:
    doc = Document()
    part, root = get_or_create_part(doc, FOOTNOTES_SPEC)
    assert isinstance(part, _FootnotesPart)
    assert part.content_type == CT.WML_FOOTNOTES
    assert part.partname == "/word/footnotes.xml"
    assert root.tag == qn("w:footnotes")
    # ECMA-376 / Word convention: separator + continuationSeparator (C1).
    note_ids = [n.get(qn("w:id")) for n in root.findall(qn("w:footnote"))]
    assert note_ids == ["-1", "0"]


def test_get_or_create_endnotes_creates_when_absent() -> None:
    doc = Document()
    part, root = get_or_create_part(doc, ENDNOTES_SPEC)
    assert isinstance(part, _EndnotesPart)
    assert part.content_type == CT.WML_ENDNOTES
    assert part.partname == "/word/endnotes.xml"
    assert root.tag == qn("w:endnotes")
    note_ids = [n.get(qn("w:id")) for n in root.findall(qn("w:endnote"))]
    assert note_ids == ["-1", "0"]


def test_create_wires_relationship_from_document_part() -> None:
    doc = Document()
    part, _ = get_or_create_part(doc, FOOTNOTES_SPEC)
    # Round-trip the relationship lookup that the helper itself performs.
    assert doc.part.part_related_by(RT.FOOTNOTES) is part


# --------------------------------------------------------------------------
# Idempotency: second call returns the same part, not a freshly built one.
# --------------------------------------------------------------------------


def test_second_call_returns_same_part() -> None:
    doc = Document()
    first, first_root = get_or_create_part(doc, COMMENTS_SPEC)
    second, second_root = get_or_create_part(doc, COMMENTS_SPEC)
    assert first is second
    assert first_root is second_root


def test_idempotency_preserves_mutations() -> None:
    """Edits between the two calls must survive the second lookup."""
    doc = Document()
    _, root = get_or_create_part(doc, FOOTNOTES_SPEC)
    sub(root, "w:footnote", **{"w:id": "1"})
    _, root_again = get_or_create_part(doc, FOOTNOTES_SPEC)
    assert root is root_again
    # 2 seeded separators (ids -1, 0) + 1 user-added footnote (id 1) = 3.
    assert len(list(root_again)) == 3
    note_ids = [n.get(qn("w:id")) for n in root_again]
    assert note_ids == ["-1", "0", "1"]


# --------------------------------------------------------------------------
# Round-trip: parts survive save/reopen with the registered class.
# --------------------------------------------------------------------------


def test_footnotes_part_round_trip(tmp_path: Path) -> None:
    doc = Document()
    _, root = get_or_create_part(doc, FOOTNOTES_SPEC)
    sub(root, "w:footnote", **{"w:id": "1"})
    out = tmp_path / "with_footnotes.docx"
    doc.save(str(out))

    reopened = Document(str(out))
    part_again, root_again = get_or_create_part(reopened, FOOTNOTES_SPEC)
    assert isinstance(part_again, _FootnotesPart)
    # The footnote we wrote AND the seeded separators (C1) survive round-trip.
    note_ids = [n.get(qn("w:id")) for n in root_again]
    assert note_ids == ["-1", "0", "1"]
    # Verify the separator types are correctly typed for Word to render the
    # horizontal divider line above the footnote area.
    types = {
        n.get(qn("w:id")): n.get(qn("w:type"))
        for n in root_again
    }
    assert types["-1"] == "separator"
    assert types["0"] == "continuationSeparator"
    assert types["1"] is None  # user note has no type attribute


def test_endnotes_part_round_trip(tmp_path: Path) -> None:
    doc = Document()
    _, root = get_or_create_part(doc, ENDNOTES_SPEC)
    sub(root, "w:endnote", **{"w:id": "1"})
    out = tmp_path / "with_endnotes.docx"
    doc.save(str(out))

    reopened = Document(str(out))
    part_again, root_again = get_or_create_part(reopened, ENDNOTES_SPEC)
    assert isinstance(part_again, _EndnotesPart)
    note_ids = [n.get(qn("w:id")) for n in root_again]
    assert note_ids == ["-1", "0", "1"]
    types = {
        n.get(qn("w:id")): n.get(qn("w:type"))
        for n in root_again
    }
    assert types["-1"] == "separator"
    assert types["0"] == "continuationSeparator"


def test_no_part_created_when_not_requested(tmp_path: Path) -> None:
    """``get_or_create_part`` must be the only thing that adds the
    relationship — a save/reopen without calling it must not have one."""
    doc = Document()
    out = tmp_path / "no_optional_parts.docx"
    doc.save(str(out))

    reopened = Document(str(out))
    for rel in reopened.part.rels.values():
        assert rel.reltype not in {RT.COMMENTS, RT.FOOTNOTES, RT.ENDNOTES}


# --------------------------------------------------------------------------
# Falls back to plain XmlPart for unknown content types.
# --------------------------------------------------------------------------


def test_unknown_content_type_falls_back_to_xmlpart() -> None:
    from docx_plus.core.parts import PartSpec

    custom = PartSpec(
        partname="/word/custom.xml",
        content_type="application/vnd.example.custom+xml",
        relationship_type="http://example.com/custom",
        root_xml=(
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            b'<w:custom xmlns:w="http://schemas.openxmlformats.org/wordprocessingml'
            b'/2006/main"/>'
        ),
    )
    doc = Document()
    part, _ = get_or_create_part(doc, custom)
    assert type(part) is XmlPart
