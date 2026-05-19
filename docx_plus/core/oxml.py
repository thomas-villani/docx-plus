"""Element-construction helpers — the single chokepoint for OOXML I/O.

Capability modules must construct elements through :func:`el` and :func:`sub`
rather than calling ``lxml.etree`` directly. SPEC §9.2.
"""

from __future__ import annotations

from typing import Any

from lxml import etree

from docx_plus.core.ns import NSMAP, qn


def _resolve_attr_key(key: str) -> str:
    """Translate a ``"prefix:local"`` attribute key to Clark notation.

    Plain keys (no ``:``) are returned unchanged so callers can mix namespaced
    and bare attributes naturally.
    """
    return qn(key) if ":" in key else key


def el(tag: str, **attrs: str) -> etree._Element:
    """Create a namespaced element with attributes.

    Args:
        tag: Element name in ``prefix:local`` form. The prefix must be a key in
            :data:`docx_plus.core.ns.NSMAP`.
        **attrs: Attributes. Keys may be namespaced (``"w:val"``) or plain
            (``"id"``). Values are coerced to strings via the lxml setter.

    Returns:
        A fresh detached :class:`lxml.etree._Element`.

    Example:
        >>> style = el("w:style", **{"w:type": "paragraph", "w:styleId": "Foo"})
        >>> style.tag.endswith("}style")
        True
    """
    node = etree.Element(qn(tag), nsmap=NSMAP)
    for key, value in attrs.items():
        node.set(_resolve_attr_key(key), value)
    return node


def sub(parent: etree._Element, tag: str, **attrs: str) -> etree._Element:
    """Create a namespaced child of ``parent`` and append it.

    Equivalent to :func:`el` followed by ``parent.append(...)``. Returned for
    chained use.
    """
    child = el(tag, **attrs)
    parent.append(child)
    return child


def xpath(node: etree._Element, expr: str, **variables: Any) -> list[Any]:
    """Run an XPath query with the library's namespace map pre-bound.

    Uses :class:`lxml.etree.XPath` so the call works equally on raw lxml
    elements and on python-docx's :class:`BaseOxmlElement` subclasses (whose
    own ``xpath`` method does not accept a ``namespaces=`` kwarg).

    Values that vary per call (style ids, numbering ids, etc.) should be
    passed as ``**variables`` and referenced with ``$name`` in the
    expression — this lets lxml escape them as XPath variables rather than
    splicing the caller's string into the query, which avoids quote-handling
    bugs and the obvious injection class.

    Args:
        node: Context node.
        expr: XPath expression that may reference ``w:``, ``w14:``, ``r:``,
            ``mc:``, ``a:`` prefixes plus any ``$name`` placeholders bound by
            ``**variables``.
        **variables: Values for ``$name`` placeholders in ``expr``. lxml
            accepts ``bool``, ``int``, ``float``, and ``str``.

    Returns:
        The XPath result list (elements, attribute strings, etc.). Returned
        as ``list[Any]`` because XPath result types vary.

    Example:
        >>> # safe variable binding instead of f-string interpolation
        >>> # xpath(styles_root, "./w:style[@w:styleId=$sid]", sid=style_id)
    """
    compiled = etree.XPath(expr, namespaces=NSMAP)
    result = compiled(node, **variables)
    if isinstance(result, list):
        return result
    return [result]


def remove(node: etree._Element) -> None:
    """Detach ``node`` from its parent if it has one. No-op if already detached.

    Args:
        node: Element to remove.
    """
    parent = node.getparent()
    if parent is not None:
        parent.remove(node)


__all__ = ["el", "remove", "sub", "xpath"]
