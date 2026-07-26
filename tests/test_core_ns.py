"""Tests for ``docx_plus.core.ns``."""

from __future__ import annotations

import pytest

from docx_plus.core.ns import BUILD_NSMAP, MC, NSMAP, W14, W15, XML, A, R, W, qn


def test_qn_main_namespace() -> None:
    assert qn("w:tag") == f"{{{W}}}tag"


def test_qn_w15_namespace() -> None:
    assert qn("w15:commentEx") == f"{{{W15}}}commentEx"


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
    assert set(NSMAP) == {"w", "w14", "w15", "r", "mc", "a", "xml"}


def test_build_nsmap_omits_extension_prefixes() -> None:
    # BUILD_NSMAP is what ``el`` declares on main-document elements. w15
    # lives only in commentsExtended.xml, so declaring it here would put a
    # stray xmlns:w15 on every element the library writes into document.xml.
    assert set(BUILD_NSMAP) == set(NSMAP) - {"w15"}
    assert all(BUILD_NSMAP[prefix] == NSMAP[prefix] for prefix in BUILD_NSMAP)
