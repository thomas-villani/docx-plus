"""Package part / relationship helpers for separate OOXML parts.

v0.1 capabilities (styles, fields, controls, protection) mutated only the
main document part and ``settings.xml`` — both already present in every
docx. v0.2 introduces capabilities backed by *separate* OOXML parts that
may not exist in a fresh document:

- ``/word/comments.xml`` (relationship :data:`RT.COMMENTS`)
- ``/word/footnotes.xml`` (relationship :data:`RT.FOOTNOTES`)
- ``/word/endnotes.xml`` (relationship :data:`RT.ENDNOTES`)
- ``/word/commentsExtended.xml`` (relationship
  :data:`RT_COMMENTS_EXTENDED`) — comment threading, added in v0.4
- ``/word/commentsIds.xml`` (relationship :data:`RT_COMMENTS_IDS`) and
  ``/word/people.xml`` (relationship :data:`RT_PEOPLE`) — durable comment
  ids and author presence, added in v0.5
- ``/word/numbering.xml`` (relationship :data:`RT.NUMBERING`) — list
  definitions, added in v0.5

This module provides a single :func:`get_or_create_part` helper that the
``comments``, ``notes``, and ``numbering`` packages use to look up an
existing part or fabricate a fresh one with an empty default root.

python-docx already registers ``CommentsPart`` with
``PartFactory.part_type_for[CT.WML_COMMENTS]`` — and ``NumberingPart``
with ``CT.WML_NUMBERING`` — so existing comments and numbering parts
deserialize with a parsed ``.element``. It does *not* register
footnote / endnote / commentsExtended / commentsIds / people part
classes, so this module installs minimal
:class:`~docx.opc.part.XmlPart` subclasses for those content types at
import time. Without that registration a round-tripped document with
existing footnotes would surface them as raw blobs.

Numbering is the one part here that :func:`get_or_create_part` must own
rather than delegating to python-docx: ``DocumentPart.numbering_part``
fabricates through ``NumberingPart.new()``, which is an unimplemented
stub that raises ``NotImplementedError``, so it works only for documents
that already have the part.

SPEC §2 lists the part / relationship plumbing as a v0.2 responsibility.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, cast

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import PartFactory, XmlPart
from docx.oxml.parser import parse_xml
from lxml import etree

if TYPE_CHECKING:
    from docx.document import Document


# ---------------------------------------------------------------------------
# Round-trip support for footnote / endnote parts.
#
# python-docx ships an explicit ``CommentsPart`` and registers it in
# ``part_type_for`` at package import. Footnote and endnote parts are not
# registered, so an existing document with footnotes would otherwise load
# them as the default ``Part`` (blob-only). Subclassing :class:`XmlPart` is
# enough to make load-time deserialization parse ``.element``.
# ---------------------------------------------------------------------------


# NOTE: ``tests/test_core_parts.py`` imports these private classes by name to
# assert ``PartFactory.part_type_for`` is wired to them. If any of them is
# renamed or inlined, update that test in the same change (L10).
class _FootnotesPart(XmlPart):
    """Internal :class:`XmlPart` subclass for ``/word/footnotes.xml``."""


class _EndnotesPart(XmlPart):
    """Internal :class:`XmlPart` subclass for ``/word/endnotes.xml``."""


class _CommentsExtendedPart(XmlPart):
    """Internal :class:`XmlPart` subclass for ``/word/commentsExtended.xml``."""


class _CommentsIdsPart(XmlPart):
    """Internal :class:`XmlPart` subclass for ``/word/commentsIds.xml``."""


class _PeoplePart(XmlPart):
    """Internal :class:`XmlPart` subclass for ``/word/people.xml``."""


# The comment side-part content / relationship types are Microsoft
# extensions introduced with Word 2013's threaded comments and Word 2016's
# durable ids. python-docx's ``CT`` and ``RT`` enums predate them and carry
# no member, so the URIs live here as plain constants.
CT_COMMENTS_EXTENDED = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
)
RT_COMMENTS_EXTENDED = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"

CT_COMMENTS_IDS = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"
RT_COMMENTS_IDS = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"

CT_PEOPLE = "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml"
RT_PEOPLE = "http://schemas.microsoft.com/office/2011/relationships/people"


PartFactory.part_type_for.setdefault(CT.WML_FOOTNOTES, _FootnotesPart)
PartFactory.part_type_for.setdefault(CT.WML_ENDNOTES, _EndnotesPart)
PartFactory.part_type_for.setdefault(CT_COMMENTS_EXTENDED, _CommentsExtendedPart)
PartFactory.part_type_for.setdefault(CT_COMMENTS_IDS, _CommentsIdsPart)
PartFactory.part_type_for.setdefault(CT_PEOPLE, _PeoplePart)


# ---------------------------------------------------------------------------
# PartSpec — identification data for an OOXML part. The three canonical
# specs below describe every separate part v0.2 cares about.
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class PartSpec:
    """Identification for an OOXML part looked up or fabricated as a unit.

    Args:
        partname: Absolute package URI, e.g. ``"/word/comments.xml"``.
        content_type: Content-type URI from
            :class:`docx.opc.constants.CONTENT_TYPE`.
        relationship_type: Relationship-type URI from
            :class:`docx.opc.constants.RELATIONSHIP_TYPE`.
        root_xml: Complete XML bytes for a fresh, empty root element of
            this part (used only when creating a part that doesn't exist).
    """

    partname: str
    content_type: str
    relationship_type: str
    root_xml: bytes


_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
_W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
_W16CID_NS = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def _empty_root(local_name: str) -> bytes:
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b"<w:" + local_name.encode() + b' xmlns:w="' + _W_NS.encode() + b'"/>'
    )


def _comments_root() -> bytes:
    """Return a fresh ``comments.xml`` root declaring the ``w14`` extension.

    Threaded comments stamp ``w14:paraId`` onto every comment body
    paragraph (see :mod:`docx_plus.comments`), so the root has to declare
    that prefix. ``mc:Ignorable="w14"`` is the Markup Compatibility
    contract that lets a consumer which does not understand the 2010
    extensions drop the attribute instead of rejecting the part — the
    same declaration Word itself writes.
    """
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b'<w:comments xmlns:w="' + _W_NS.encode() + b'"'
        b' xmlns:w14="' + _W14_NS.encode() + b'"'
        b' xmlns:mc="' + _MC_NS.encode() + b'"'
        b' mc:Ignorable="w14"/>'
    )


def _comments_extended_root() -> bytes:
    """Return a fresh, empty ``commentsExtended.xml`` root."""
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b'<w15:commentsEx xmlns:w15="' + _W15_NS.encode() + b'"/>'
    )


def _comments_ids_root() -> bytes:
    """Return a fresh, empty ``commentsIds.xml`` root."""
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b'<w16cid:commentsIds xmlns:w16cid="' + _W16CID_NS.encode() + b'"/>'
    )


def _people_root() -> bytes:
    """Return a fresh, empty ``people.xml`` root."""
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b'<w15:people xmlns:w15="' + _W15_NS.encode() + b'"/>'
    )


def _notes_root_with_separators(root_tag: str, child_tag: str, marker_tag: str) -> bytes:
    """Return a footnotes / endnotes root pre-seeded with separator entries.

    ECMA-376 17.11.16 / 17.11.20 ("Footnote / Endnote Separator") and
    17.11.7 ("Continuation Separator") describe two reserved entries Word
    expects in every footnotes.xml / endnotes.xml: ``w:id="-1"`` with
    ``w:type="separator"`` (the horizontal line between body text and the
    first footnote on a page) and ``w:id="0"`` with
    ``w:type="continuationSeparator"`` (the line for footnotes that span
    pages). Omitting them produces files that Word may surface "needs
    repair" prompts for, and that strict consumers may reject. The ids
    -1 and 0 are reserved (see :class:`~docx_plus.notes.registry`); user
    notes start at 1.
    """
    enc = root_tag.encode()
    child = child_tag.encode()
    marker = marker_tag.encode()
    cont = b"continuationSeparator"
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b"<w:" + enc + b' xmlns:w="' + _W_NS.encode() + b'">'
        b"<w:" + child + b' w:id="-1" w:type="separator">'
        b"<w:p><w:r><w:" + marker + b"/></w:r></w:p>"
        b"</w:" + child + b">"
        b"<w:" + child + b' w:id="0" w:type="continuationSeparator">'
        b"<w:p><w:r><w:" + cont + b"/></w:r></w:p>"
        b"</w:" + child + b">"
        b"</w:" + enc + b">"
    )


COMMENTS_SPEC = PartSpec(
    partname="/word/comments.xml",
    content_type=CT.WML_COMMENTS,
    relationship_type=RT.COMMENTS,
    root_xml=_comments_root(),
)

COMMENTS_EXTENDED_SPEC = PartSpec(
    partname="/word/commentsExtended.xml",
    content_type=CT_COMMENTS_EXTENDED,
    relationship_type=RT_COMMENTS_EXTENDED,
    root_xml=_comments_extended_root(),
)

COMMENTS_IDS_SPEC = PartSpec(
    partname="/word/commentsIds.xml",
    content_type=CT_COMMENTS_IDS,
    relationship_type=RT_COMMENTS_IDS,
    root_xml=_comments_ids_root(),
)

PEOPLE_SPEC = PartSpec(
    partname="/word/people.xml",
    content_type=CT_PEOPLE,
    relationship_type=RT_PEOPLE,
    root_xml=_people_root(),
)

#: ``/word/numbering.xml`` — list definitions.
#:
#: Unlike the comment side-parts this needs no ``PartFactory``
#: registration: python-docx registers its own ``NumberingPart`` for
#: :data:`CT.WML_NUMBERING`, and :func:`get_or_create_part` looks the
#: class up, so an existing part round-trips as parsed XML for free. What
#: python-docx *cannot* do is fabricate a missing one —
#: ``NumberingPart.new()`` raises ``NotImplementedError`` — which is
#: exactly the gap this spec closes.
#:
#: An empty ``<w:numbering/>`` is schema-valid: every child of
#: ``CT_Numbering`` is optional, and there are no reserved entries to
#: pre-seed the way footnotes and endnotes need separators.
NUMBERING_SPEC = PartSpec(
    partname="/word/numbering.xml",
    content_type=CT.WML_NUMBERING,
    relationship_type=RT.NUMBERING,
    root_xml=_empty_root("numbering"),
)

FOOTNOTES_SPEC = PartSpec(
    partname="/word/footnotes.xml",
    content_type=CT.WML_FOOTNOTES,
    relationship_type=RT.FOOTNOTES,
    root_xml=_notes_root_with_separators("footnotes", "footnote", "separator"),
)

ENDNOTES_SPEC = PartSpec(
    partname="/word/endnotes.xml",
    content_type=CT.WML_ENDNOTES,
    relationship_type=RT.ENDNOTES,
    root_xml=_notes_root_with_separators("endnotes", "endnote", "separator"),
)


# ---------------------------------------------------------------------------
# get_or_create_part — single entry point for the comments / notes modules.
# ---------------------------------------------------------------------------


def get_or_create_part(doc: Document, spec: PartSpec) -> tuple[XmlPart, etree._Element]:
    """Return ``(part, root_element)`` for the part identified by ``spec``.

    If the main document part already has a relationship of
    ``spec.relationship_type``, returns the related part. Otherwise
    creates a fresh part with the empty default root from ``spec.root_xml``
    and wires the relationship from the document part. Idempotent — a
    second call with the same spec returns the same part.

    The part class is looked up in :attr:`PartFactory.part_type_for` so
    callers get the most specific class registered for the content type
    (``CommentsPart`` for comments; the internal
    :class:`_FootnotesPart` / :class:`_EndnotesPart` for the note parts).
    Unknown content types fall back to :class:`XmlPart`.

    Args:
        doc: The python-docx :class:`~docx.document.Document` whose part
            graph is being mutated.
        spec: Identifier for the target part. Use :data:`COMMENTS_SPEC`,
            :data:`FOOTNOTES_SPEC`, :data:`ENDNOTES_SPEC`, or build a
            :class:`PartSpec` for a custom part.

    Returns:
        A tuple ``(part, root_element)``. ``root_element`` is
        ``part.element`` — the XML root that caller modules mutate.
    """
    document_part = doc.part
    try:
        part = cast("XmlPart", document_part.part_related_by(spec.relationship_type))
    except KeyError:
        package = document_part.package
        element = parse_xml(spec.root_xml)
        part_cls = cast(
            "type[XmlPart]",
            PartFactory.part_type_for.get(spec.content_type, XmlPart),
        )
        part = part_cls(PackURI(spec.partname), spec.content_type, element, package)
        document_part.relate_to(part, spec.relationship_type)
    return part, part.element


__all__ = [
    "COMMENTS_EXTENDED_SPEC",
    "COMMENTS_IDS_SPEC",
    "COMMENTS_SPEC",
    "CT_COMMENTS_EXTENDED",
    "CT_COMMENTS_IDS",
    "CT_PEOPLE",
    "ENDNOTES_SPEC",
    "FOOTNOTES_SPEC",
    "NUMBERING_SPEC",
    "PEOPLE_SPEC",
    "PartSpec",
    "RT_COMMENTS_EXTENDED",
    "RT_COMMENTS_IDS",
    "RT_PEOPLE",
    "get_or_create_part",
]
