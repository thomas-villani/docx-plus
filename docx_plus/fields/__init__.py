"""Field insertion and update helpers — SPEC §7.

Public API:

* :func:`add_page_number_field` — PAGE / NUMPAGES / SECTIONPAGES
* :func:`add_date_field` — DATE / CREATEDATE
* :func:`add_field` — generic complex field
* :func:`read_fields` — read the fields already in a document
* :func:`mark_fields_dirty` — flag ``w:updateFields`` in settings.xml
"""

from __future__ import annotations

from docx_plus.fields.read import FieldInfo, read_fields
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
    "FieldInfo",
    "PageFieldName",
    "StyleRefNumber",
    "add_date_field",
    "add_field",
    "add_page_number_field",
    "add_style_reference",
    "mark_fields_dirty",
    "read_fields",
]
