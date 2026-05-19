"""Page-layout extras beyond python-docx.

Columns, mid-document section breaks, and doc-level distinct even/odd
headers.

python-docx already exposes orientation, margins, page size, per-section
header/footer (including first-page variants), section start type, and
``Document.add_section()`` (append-only). This module fills three
documented gaps:

- :func:`set_columns` — ``<w:cols>`` is not abstracted by python-docx.
- :func:`insert_section_break` — ``add_section`` only appends; inserting
  mid-document requires moving the trailing ``<w:sectPr>``.
- :func:`enable_distinct_even_odd_headers` /
  :func:`disable_distinct_even_odd_headers` — the doc-level
  ``<w:evenAndOddHeaders/>`` flag in ``settings.xml`` is distinct from
  the per-section ``Section.different_first_page_header_footer``
  python-docx already exposes, and is constantly confused with it.

See SPEC §15 (deferred to v0.2) and ``notes-v0_1-scope.md §2.1`` for
context.
"""

from __future__ import annotations

from docx_plus.layout.breaks import SectionStartType, insert_section_break
from docx_plus.layout.columns import set_columns
from docx_plus.layout.settings import (
    disable_distinct_even_odd_headers,
    enable_distinct_even_odd_headers,
)

__all__ = [
    "SectionStartType",
    "disable_distinct_even_odd_headers",
    "enable_distinct_even_odd_headers",
    "insert_section_break",
    "set_columns",
]
