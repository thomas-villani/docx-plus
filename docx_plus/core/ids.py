"""Per-document registry of issued SDT ``w:id`` values.

Only ``w:id`` on ``w:sdt`` elements is in scope for v0.1. Other ID-like
attributes (``r:id``, bookmark ``w:id``, comment ``w:id``) live in separate
namespaces with separate uniqueness requirements and will use their own
registries when they're needed. SPEC §3, IMPLEMENTATION.md §7.
"""

from __future__ import annotations

import secrets
from typing import TYPE_CHECKING

from lxml import etree

from docx_plus.core import DocxPlusError
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


class IdRegistry:
    """Tracks issued ``w:id`` values for one document-edit session.

    On construction, the registry scans the document body and settings part
    for existing ``w:id`` values on ``w:sdt`` descendants and seeds itself
    with them, so :meth:`next` cannot collide with values already in the file.

    Lifecycle: one registry per document. Pass it explicitly to functions
    that need IDs; do not attach it as a magic attribute on the
    :class:`docx.document.Document` (SPEC §9.4).
    """

    def __init__(self, doc: Document) -> None:
        """Scan ``doc`` for existing SDT IDs.

        Args:
            doc: A python-docx :class:`~docx.document.Document`.
        """
        self._issued: set[int] = set()
        self._seed_from_document(doc)

    def _seed_from_document(self, doc: Document) -> None:
        body_element = doc.element.body
        self._collect_sdt_ids(body_element)

        settings_part = getattr(doc, "settings", None)
        settings_element = getattr(settings_part, "element", None)
        if settings_element is not None:
            self._collect_sdt_ids(settings_element)

    def _collect_sdt_ids(self, root: etree._Element) -> None:
        for id_el in xpath(root, ".//w:sdt/w:sdtPr/w:id"):
            raw = id_el.get(qn("w:val"))
            if raw is None:
                continue
            try:
                value = int(raw)
            except ValueError:
                continue
            self._issued.add(value)

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
        raise RuntimeError("IdRegistry exhausted the 31-bit ID space")

    def reserve(self, value: int) -> int:
        """Reserve a specific value, asserting it isn't already issued.

        Args:
            value: A positive integer in ``[1, 2**31 - 1]``.

        Returns:
            ``value`` (echoed so the call composes inline).

        Raises:
            DuplicateIdError: If ``value`` has already been issued or
                reserved on this registry.
            ValueError: If ``value`` is outside the 31-bit positive range.
        """
        if not 1 <= value <= _MAX_W_ID:
            raise ValueError(f"id {value!r} outside 31-bit positive range")
        if value in self._issued:
            raise DuplicateIdError(f"id {value} already issued")
        self._issued.add(value)
        return value

    def issued(self) -> frozenset[int]:
        """Return an immutable snapshot of all issued ids."""
        return frozenset(self._issued)


__all__ = ["DuplicateIdError", "IdRegistry"]
