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
    add_date_field,
    add_field,
    add_page_number_field,
)
from docx_plus.fields.update import mark_fields_dirty

__all__ = [
    "PageFieldName",
    "add_date_field",
    "add_field",
    "add_page_number_field",
    "mark_fields_dirty",
]
