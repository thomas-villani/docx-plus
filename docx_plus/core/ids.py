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

SPEC §3, IMPLEMENTATION.md §7.
"""

from __future__ import annotations

import secrets
from typing import TYPE_CHECKING

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


__all__ = ["DuplicateIdError", "IdRangeError", "IdRegistry"]
