"""Tests for ``docx_plus.core.ns``."""

from __future__ import annotations

import pytest

from docx_plus.core.ns import MC, NSMAP, W14, XML, A, R, W, qn


def test_qn_main_namespace() -> None:
    assert qn("w:tag") == f"{{{W}}}tag"


def test_qn_w14_namespace() -> None:
    assert qn("w14:checkbox") == f"{{{W14}}}checkbox"


def test_qn_relationships_namespace() -> None:
    assert qn("r:id") == f"{{{R}}}id"


def test_qn_drawing_namespace() -> None:
    assert qn("a:srgbClr") == f"{{{A}}}srgbClr"


def test_qn_markup_compatibility_namespace() -> None:
    assert qn("mc:Choice") == f"{{{MC}}}Choice"


def test_qn_xml_namespace() -> None:
    """``xml:space`` is needed by ``w:instrText`` to preserve field whitespace."""
    assert qn("xml:space") == f"{{{XML}}}space"


def test_qn_rejects_unqualified() -> None:
    with pytest.raises(ValueError, match="prefix:local"):
        qn("notqualified")


def test_qn_rejects_unknown_prefix() -> None:
    with pytest.raises(ValueError, match="unknown namespace prefix"):
        qn("xyzzy:thing")


def test_nsmap_keys() -> None:
    assert set(NSMAP) == {"w", "w14", "r", "mc", "a", "xml"}
