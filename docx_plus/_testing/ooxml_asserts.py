"""Shared OOXML assertion helpers used across the test suite.

Internal API — not part of the public surface. Built out lazily as later
phases introduce call sites (SPEC §10).
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx_plus.core.ns import qn
from docx_plus.core.oxml import xpath

if TYPE_CHECKING:
    from docx.document import Document


def assert_ids_unique(doc: Document) -> None:
    """Assert every ``w:id`` on a ``w:sdt`` descendant is unique within the doc.

    Args:
        doc: python-docx Document to inspect.

    Raises:
        AssertionError: If any ID appears more than once.
    """
    seen: dict[int, int] = {}
    for id_el in xpath(doc.element.body, ".//w:sdt/w:sdtPr/w:id"):
        raw = id_el.get(qn("w:val"))
        if raw is None:
            continue
        try:
            value = int(raw)
        except ValueError:
            continue
        seen[value] = seen.get(value, 0) + 1
    duplicates = {v: count for v, count in seen.items() if count > 1}
    assert not duplicates, f"duplicate SDT w:id values: {duplicates}"


def assert_style_defined(doc: Document, style_id: str) -> None:
    """Assert ``style_id`` is materialized in ``word/styles.xml``.

    Args:
        doc: python-docx Document to inspect.
        style_id: The style's ``w:styleId`` attribute value.

    Raises:
        AssertionError: If no ``w:style`` element with that id exists.
    """
    styles_element = doc.styles.element
    matches = xpath(styles_element, f".//w:style[@w:styleId='{style_id}']")
    assert matches, f"style {style_id!r} not defined in styles.xml"


__all__ = ["assert_ids_unique", "assert_style_defined"]
