"""Tests for ``docx_plus.core.oxml``."""

from __future__ import annotations

import pytest
from docx import Document

from docx_plus.core.ns import W, qn
from docx_plus.core.oxml import (
    _compile_xpath,
    body_document_for,
    el,
    remove,
    sub,
    xpath,
)


def test_el_creates_clark_tag() -> None:
    node = el("w:style")
    assert node.tag == f"{{{W}}}style"


def test_el_translates_namespaced_attribute_keys() -> None:
    node = el("w:style", **{"w:type": "paragraph", "w:styleId": "Foo"})
    assert node.get(qn("w:type")) == "paragraph"
    assert node.get(qn("w:styleId")) == "Foo"


def test_el_keeps_plain_attribute_keys() -> None:
    node = el("w:tag", id="bare")
    assert node.get("id") == "bare"
    assert node.get(qn("w:id")) is None


def test_sub_creates_and_appends() -> None:
    parent = el("w:styles")
    child = sub(parent, "w:style", **{"w:styleId": "Foo"})
    assert child in list(parent)
    assert child.getparent() is parent


def test_xpath_returns_matching_elements() -> None:
    parent = el("w:styles")
    sub(parent, "w:style", **{"w:styleId": "A"})
    sub(parent, "w:style", **{"w:styleId": "B"})
    sub(parent, "w:other")
    matches = xpath(parent, "./w:style")
    assert len(matches) == 2
    assert all(m.tag == f"{{{W}}}style" for m in matches)


def test_xpath_with_attribute_predicate() -> None:
    parent = el("w:styles")
    sub(parent, "w:style", **{"w:styleId": "A"})
    sub(parent, "w:style", **{"w:styleId": "B"})
    [match] = xpath(parent, "./w:style[@w:styleId='B']")
    assert match.get(qn("w:styleId")) == "B"


def test_remove_detaches_child() -> None:
    parent = el("w:styles")
    child = sub(parent, "w:style")
    assert child in list(parent)
    remove(child)
    assert child not in list(parent)
    assert child.getparent() is None


def test_remove_on_detached_node_is_noop() -> None:
    orphan = el("w:style")
    remove(orphan)
    assert orphan.getparent() is None


# --------------------------------------------------------------------------
# L11: xpath compiles each distinct expression once and caches it.
# --------------------------------------------------------------------------


def test_xpath_caches_compiled_expression() -> None:
    parent = el("w:styles")
    sub(parent, "w:style")
    first = _compile_xpath("./w:style")
    second = _compile_xpath("./w:style")
    assert first is second  # same compiled object reused


# --------------------------------------------------------------------------
# N4: body_document_for — shared proxy -> Document resolver.
# --------------------------------------------------------------------------


def test_body_document_for_returns_owning_document() -> None:
    doc = Document()
    p = doc.add_paragraph("x")
    # python-docx's DocumentPart.document builds a fresh proxy each call, so
    # compare the underlying element rather than proxy identity.
    assert body_document_for(p).element is doc.element


def test_body_document_for_rejects_non_body_proxy() -> None:
    class _FakePart:
        pass  # no .document attribute -> not the main body

    class _FakeProxy:
        part = _FakePart()

    with pytest.raises(ValueError, match="myop only supports the main document body"):
        body_document_for(_FakeProxy(), operation="myop")
