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


def build_complex_field(
    p_element: etree._Element,
    instruction: str,
    initial_text: str,
) -> etree._Element:
    r"""Append the 5-run complex-field sequence to ``p_element``.

    The sequence is::

        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve">INSTRUCTION</w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t xml:space="preserve">INITIAL_TEXT</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>

    ``xml:space="preserve"`` on the instruction and the result text
    keeps leading / trailing whitespace from being normalised by Word's
    XML reader.

    Args:
        p_element: The underlying ``w:p`` element to append to.
        instruction: The field instruction text (e.g. ``" PAGE "``,
            ``" REF bookmark1 \h "``). Surrounding spaces are part of
            the standard syntax.
        initial_text: The result text shown before Word recalculates the
            field. Use ``""`` if Word will fill it on open.

    Returns:
        The begin ``w:r`` run element — the marker for the start of the
        field.
    """
    begin_run = sub(p_element, "w:r")
    sub(begin_run, "w:fldChar", **{"w:fldCharType": "begin"})

    instr_run = sub(p_element, "w:r")
    instr_t = sub(instr_run, "w:instrText", **{"xml:space": "preserve"})
    instr_t.text = instruction

    sep_run = sub(p_element, "w:r")
    sub(sep_run, "w:fldChar", **{"w:fldCharType": "separate"})

    text_run = sub(p_element, "w:r")
    text_t = sub(text_run, "w:t", **{"xml:space": "preserve"})
    text_t.text = initial_text

    end_run = sub(p_element, "w:r")
    sub(end_run, "w:fldChar", **{"w:fldCharType": "end"})

    return begin_run


def insert_before_first_anchor(
    parent: etree._Element,
    new_element: etree._Element,
    anchor_tags: tuple[str, ...],
) -> None:
    """Insert ``new_element`` before the first ``anchor_tags`` match in ``parent``.

    Falls back to appending at the end if none of the anchors exist. The
    pattern keeps schema-strict child ordering even when ``parent`` has a
    sparse / partial set of children — most real ``settings.xml`` files
    do, so capability modules building optional ``settings.xml`` children
    use this to land in the right schema position.

    Args:
        parent: The element to insert into.
        new_element: The element to insert.
        anchor_tags: Sequence of ``prefix:local`` tag names; the first
            anchor found becomes the insertion point and ``new_element``
            is placed immediately before it. Order in the tuple should
            match schema order so the search picks the *first* later
            sibling that actually exists.

    Example:
        >>> # See `fields/update.py` and `layout/settings.py` for live uses
        >>> # against `settings.xml`.
    """
    from docx_plus.core.ns import qn as _qn

    for tag in anchor_tags:
        anchor = parent.find(_qn(tag))
        if anchor is not None:
            anchor.addprevious(new_element)
            return
    parent.append(new_element)


__all__ = [
    "build_complex_field",
    "el",
    "insert_before_first_anchor",
    "remove",
    "sub",
    "xpath",
]
