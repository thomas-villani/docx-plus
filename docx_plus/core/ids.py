"""Per-document registries of issued ``w:id`` values.

OOXML uses ``w:id`` for several disjoint namespaces — SDT controls,
bookmarks, comments, footnotes, endnotes. Each namespace has its own
uniqueness requirement; bookmark id ``7`` does not collide with comment
id ``7``. v0.1 only minted SDT ids and shipped :class:`IdRegistry` for
that purpose. v0.2 adds further namespaces (comments, bookmarks, notes)
and refactors the shared ``next``/``reserve``/``issued`` mechanics into
:class:`_IdRegistryBase`. Each namespace-specific registry is a tiny
subclass that overrides :meth:`_seed_from_document` with the right
discovery query.

:class:`ParaIdRegistry` is the one registry not backed by ``w:id``: it
mints ``w14:paraId`` values, which are hex-rendered and unique across the
whole *package* rather than within a single part. It reuses the same
31-bit allocator because that happens to be exactly the range Word
accepts for a ``paraId``.

SPEC §3, IMPLEMENTATION.md §7.
"""

from __future__ import annotations

import secrets
from typing import TYPE_CHECKING, cast

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import XmlPart
from lxml import etree

from docx_plus.core.errors import DocxPlusError
from docx_plus.core.ns import qn
from docx_plus.core.oxml import xpath

if TYPE_CHECKING:
    from docx.document import Document

_MAX_W_ID = 2**31 - 1


class DuplicateIdError(DocxPlusError, ValueError):
    """Raised when an ID is reserved twice.

    Subclasses ``ValueError`` so existing ``except ValueError:`` clauses still
    catch it; also subclasses :class:`DocxPlusError` per SPEC §9.7.
    """


class IdRangeError(DocxPlusError, ValueError):
    """Raised when a reserved ID falls outside the 31-bit positive range.

    Subclasses ``ValueError`` for backward compatibility; also subclasses
    :class:`DocxPlusError` per SPEC §9.7.
    """


class _IdRegistryBase:
    """Generic ``w:id`` tracker for a single namespace within one document.

    Subclasses customise :meth:`_seed_from_document` to discover IDs that
    already exist in the namespace they manage. Everything else
    (``next``/``reserve``/``issued``) is the same shape across namespaces
    so it lives here.

    Lifecycle: one instance per document-edit session per namespace. Pass
    the registry explicitly to functions that need IDs; do not attach
    it as a magic attribute on :class:`~docx.document.Document` (SPEC §9.4).
    """

    def __init__(self, doc: Document) -> None:
        """Scan ``doc`` for IDs already issued in this namespace.

        Args:
            doc: A python-docx :class:`~docx.document.Document`.
        """
        self._issued: set[int] = set()
        self._seed_from_document(doc)

    def _seed_from_document(self, doc: Document) -> None:  # pragma: no cover
        raise NotImplementedError

    def _collect_ids(self, root: etree._Element, expr: str) -> None:
        """Add every parseable integer ``@w:val`` returned by ``expr``.

        Helper for subclass seeders. Skips IDs that aren't integers
        (alphanumeric ``w:val`` is legal in some contexts and we don't
        try to coerce).
        """
        for id_el in xpath(root, expr):
            raw = id_el.get(qn("w:val"))
            if raw is None:
                continue
            try:
                self._issued.add(int(raw))
            except ValueError:
                continue

    def _collect_id_attrs(self, root: etree._Element, expr: str) -> None:
        """Like :meth:`_collect_ids` but for direct ``w:id`` *attributes*.

        Bookmark / comment / note range markers store the id on
        ``@w:id`` rather than as a child ``<w:id w:val="..."/>``.
        """
        for elem in xpath(root, expr):
            raw = elem.get(qn("w:id"))
            if raw is None:
                continue
            try:
                self._issued.add(int(raw))
            except ValueError:
                continue

    def _collect_hex_id_attrs(self, root: etree._Element, expr: str, attr: str) -> None:
        """Base-16 sibling of :meth:`_collect_id_attrs`.

        ``w14:paraId`` and friends are 8-hex-digit strings rather than
        decimal integers, but they share the same 31-bit uniqueness
        space, so they seed the same ``_issued`` set once parsed.
        """
        qattr = qn(attr)
        for elem in xpath(root, expr):
            raw = elem.get(qattr)
            if raw is None:
                continue
            try:
                self._issued.add(int(raw, 16))
            except ValueError:
                continue

    def next(self) -> int:
        """Issue a fresh 31-bit positive integer not previously seen.

        Returns:
            A new ``int`` in ``[1, 2**31 - 1]``.

        Raises:
            RuntimeError: If the 31-bit space is exhausted (effectively
                impossible — included for completeness).
        """
        for _ in range(64):
            candidate = secrets.randbelow(_MAX_W_ID) + 1
            if candidate not in self._issued:
                self._issued.add(candidate)
                return candidate
        for candidate in range(1, _MAX_W_ID + 1):
            if candidate not in self._issued:
                self._issued.add(candidate)
                return candidate
        raise RuntimeError("ID registry exhausted the 31-bit ID space")

    def reserve(self, value: int) -> int:
        """Reserve a specific value, asserting it isn't already issued.

        Args:
            value: A positive integer in ``[1, 2**31 - 1]``.

        Returns:
            ``value`` (echoed so the call composes inline).

        Raises:
            IdRangeError: If ``value`` is outside the 31-bit positive range.
            DuplicateIdError: If ``value`` has already been issued or
                reserved on this registry.
        """
        if not 1 <= value <= _MAX_W_ID:
            raise IdRangeError(f"id {value!r} outside 31-bit positive range")
        if value in self._issued:
            raise DuplicateIdError(f"id {value} already issued")
        self._issued.add(value)
        return value

    def issued(self) -> frozenset[int]:
        """Return an immutable snapshot of all issued ids."""
        return frozenset(self._issued)


class IdRegistry(_IdRegistryBase):
    """Tracks issued SDT ``w:id`` values for one document-edit session.

    On construction, the registry scans the document body and settings part
    for existing ``w:id`` values on ``w:sdt`` descendants and seeds itself
    with them, so :meth:`next` cannot collide with values already in the file.
    """

    def _seed_from_document(self, doc: Document) -> None:
        body_element = doc.element.body
        self._collect_ids(body_element, ".//w:sdt/w:sdtPr/w:id")

        settings_part = getattr(doc, "settings", None)
        settings_element = getattr(settings_part, "element", None)
        if settings_element is not None:
            self._collect_ids(settings_element, ".//w:sdt/w:sdtPr/w:id")


class ParaIdRegistry(_IdRegistryBase):
    """Tracks issued ``w14:paraId`` values for one document-edit session.

    ``w14:paraId`` identifies a paragraph across the whole package rather
    than within one part: threaded comments key their parent/child links
    off the ``paraId`` of a comment body's last paragraph
    (``w15:commentEx``), so a collision between a body paragraph and a
    comment paragraph would corrupt the thread graph. The registry
    therefore seeds from the document body *and* every part that can
    carry paragraphs with a ``paraId`` — comments, footnotes, endnotes.

    Word writes ``paraId`` as 8 uppercase hex digits and treats
    ``00000000`` and anything at or above ``0x80000000`` as invalid. That
    is exactly the ``[1, 2**31 - 1]`` range :meth:`next` already mints
    from, so this subclass only adds the hex rendering
    (:meth:`next_hex`) and the seeding query.
    """

    _SEED_PART_RELATIONSHIPS = (RT.COMMENTS, RT.FOOTNOTES, RT.ENDNOTES)

    def _seed_from_document(self, doc: Document) -> None:
        self._collect_hex_id_attrs(doc.element.body, ".//w:p", "w14:paraId")

        document_part = doc.part
        for reltype in self._SEED_PART_RELATIONSHIPS:
            try:
                part = cast("XmlPart", document_part.part_related_by(reltype))
            except KeyError:
                continue
            self._collect_hex_id_attrs(part.element, ".//w:p", "w14:paraId")

    def next_hex(self) -> str:
        """Issue a fresh ``paraId`` as 8 uppercase hex digits.

        Returns:
            A new ``paraId`` string such as ``"3F2A19C4"``, guaranteed
            distinct from every value seen on this registry.
        """
        return f"{self.next():08X}"


__all__ = ["DuplicateIdError", "IdRangeError", "IdRegistry", "ParaIdRegistry"]
