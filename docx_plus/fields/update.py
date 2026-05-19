"""Mark fields dirty so Word recalculates them on next open.

A complex field stores both an instruction (``w:instrText``) and a cached
result (``w:t``). Word will display the cached result unless it is told to
recalculate. Setting ``<w:updateFields w:val="true"/>`` in ``settings.xml``
flips a one-shot flag — Word resolves every field in the document on open,
then clears the flag back to ``false`` so the recalculation does not repeat.

This module imports only from ``docx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from docx_plus.core.ns import qn
from docx_plus.core.oxml import el

if TYPE_CHECKING:
    from docx.document import Document as DocxDocument


# Schema-strict insertion: in ``CT_Settings`` (ECMA-376 17.15.1.78),
# ``w:updateFields`` sits in the late group between ``w:alwaysMergeEmptyNamespace``
# and ``w:hdrShapeDefaults``. We insert before the first of these anchor
# children we find; if none exist we fall back to appending.
_UPDATE_FIELDS_LATER_SIBLINGS: tuple[str, ...] = (
    "w:hdrShapeDefaults",
    "w:footnotePr",
    "w:endnotePr",
    "w:compat",
    "w:docVars",
    "w:rsids",
    "w:mathPr",
    "w:themeFontLang",
    "w:clrSchemeMapping",
    "w:shapeDefaults",
    "w:decimalSymbol",
    "w:listSeparator",
)


def _insert_before_first_anchor(
    parent: etree._Element,
    new_element: etree._Element,
    anchor_tags: tuple[str, ...],
) -> None:
    """Insert ``new_element`` before the first ``anchor_tags`` match in ``parent``.

    Falls back to appending at the end if none of the anchors exist. The
    pattern keeps schema-strict child ordering even when ``parent`` has a
    sparse / partial set of children (which most real-world ``settings.xml``
    files do).
    """
    for tag in anchor_tags:
        anchor = parent.find(qn(tag))
        if anchor is not None:
            anchor.addprevious(new_element)
            return
    parent.append(new_element)


def mark_fields_dirty(doc: DocxDocument) -> None:
    """Set ``w:updateFields="true"`` in ``settings.xml``.

    Idempotent: calling twice produces a single ``w:updateFields`` element.
    If an existing element has ``w:val="false"``, it is updated to ``"true"``.

    Args:
        doc: The python-docx :class:`~docx.document.Document` whose settings
            part should be flagged.

    Example:
        >>> from docx import Document
        >>> from docx_plus.fields import add_page_number_field, mark_fields_dirty
        >>> doc = Document()
        >>> p = doc.add_paragraph("Page ")
        >>> _ = add_page_number_field(p)
        >>> mark_fields_dirty(doc)
    """
    settings = doc.settings.element
    existing = settings.find(qn("w:updateFields"))
    if existing is not None:
        existing.set(qn("w:val"), "true")
        return
    new = el("w:updateFields", **{"w:val": "true"})
    _insert_before_first_anchor(settings, new, _UPDATE_FIELDS_LATER_SIBLINGS)


__all__ = ["mark_fields_dirty"]
