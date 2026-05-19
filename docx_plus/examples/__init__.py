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
"""
