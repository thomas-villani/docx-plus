"""Tests for ``docx_plus.core.oxml``."""

from __future__ import annotations

from docx_plus.core.ns import W, qn
from docx_plus.core.oxml import el, remove, sub, xpath


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
