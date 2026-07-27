"""Field insertion and update helpers — SPEC §7.

Public API:

* :func:`add_page_number_field` — PAGE / NUMPAGES / SECTIONPAGES
* :func:`add_date_field` — DATE / CREATEDATE
* :func:`add_field` — generic complex field
* :func:`mark_fields_dirty` — flag ``w:updateFields`` in settings.xml
"""

from __future__ import annotations

from docx_plus.fields.simple import (
    PageFieldName,
    StyleRefNumber,
    add_date_field,
    add_field,
    add_page_number_field,
    add_style_reference,
)
from docx_plus.fields.update import mark_fields_dirty

__all__ = [
    "PageFieldName",
    "StyleRefNumber",
    "add_date_field",
    "add_field",
    "add_page_number_field",
    "add_style_reference",
    "mark_fields_dirty",
]
