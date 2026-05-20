r"""Long-document publishing primitives — TOC, captions, table of figures.

The publishing module composes existing field plumbing into the
primitives that make Word a viable long-document tool: a top-level
Table of Contents, captioned figures / tables, and a downstream Table
of Figures that picks them up. Each helper emits a complex field;
Word populates the visible result on next open (so callers should run
:func:`docx_plus.fields.mark_fields_dirty` before saving — see
docstrings for the exact pattern).

Public surface:

- :func:`add_toc` — ``TOC`` complex field
- :func:`add_caption` — ``SEQ`` complex field with a leading label run
- :func:`add_table_of_figures` — ``TOC \c "Figure"`` complex field

Bibliography (sources + citations + ``BIBLIOGRAPHY`` field) is deferred
to v0.3 because it rides on the Custom XML Parts data-binding
subsystem (also v0.3).

See SPEC §15 (the post-v0.1 roadmap).
"""

from __future__ import annotations

from docx_plus.publishing.captions import add_caption
from docx_plus.publishing.figures import add_table_of_figures
from docx_plus.publishing.toc import add_toc

__all__ = [
    "add_caption",
    "add_table_of_figures",
    "add_toc",
]
