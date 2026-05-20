"""Runnable example scripts demonstrating the docx_plus surface.

Each module is a standalone executable: ``python -m
docx_plus.examples.<name>``. The smoke test in
``tests/test_examples_smoke.py`` runs every one with no arguments to catch
public-API drift.

Modules:

- :mod:`docx_plus.examples.inspect_document` — print effective formatting +
  provenance for each paragraph (SPEC §11).
- :mod:`docx_plus.examples.restyle_existing` — change Heading1, save, prove
  the cascade reflects it.
- :mod:`docx_plus.examples.build_form` — every Phase 4 control type plus
  ``protect_document(mode="forms")``.
- :mod:`docx_plus.examples.populate_form` — read/set/clear round-trip on
  the form ``build_form`` produces.
- :mod:`docx_plus.examples.add_comments` — anchored comments (v0.2).
- :mod:`docx_plus.examples.multi_column_layout` — columns, mid-doc section
  breaks, distinct even/odd headers (v0.2).
- :mod:`docx_plus.examples.bookmarks_and_xrefs` — bookmarks plus REF /
  PAGEREF cross-references (v0.2).
- :mod:`docx_plus.examples.footnotes_and_endnotes` — footnotes and
  endnotes (v0.2).
- :mod:`docx_plus.examples.publishing_layout` — TOC + captions + Table
  of Figures (v0.2 expansion).
"""
