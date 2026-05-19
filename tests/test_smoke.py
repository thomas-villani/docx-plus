"""Trivial smoke test — imports the package and asserts the error base exists."""

from __future__ import annotations

import docx_plus


def test_package_importable() -> None:
    assert docx_plus.__version__ == "0.2.0"
    assert issubclass(docx_plus.DocxPlusError, Exception)
