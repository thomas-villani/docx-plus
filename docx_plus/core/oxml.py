"""Element-construction helpers — the single chokepoint for OOXML I/O.

Capability modules must construct elements through :func:`el` and :func:`sub`
rather than calling ``lxml.etree`` directly. SPEC §9.2.
"""

from __future__ import annotations

import re
from functools import lru_cache
from typing import TYPE_CHECKING, Any, cast

from lxml import etree

from docx_plus.core.ns import BUILD_NSMAP, NSMAP, qn

if TYPE_CHECKING:
    from docx.document import Document


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

    Elements in the main document namespaces are built declaring
    :data:`~docx_plus.core.ns.BUILD_NSMAP`. An element in any other
    namespace — ``w15`` and ``w16cid``, which appear only in the comment
    side-parts — declares just its own prefix, so extension parts stay
    free of irrelevant declarations and ``document.xml`` stays free of
    the extension prefixes.

    Returns:
        A fresh detached :class:`lxml.etree._Element`.

    Example:
        >>> style = el("w:style", **{"w:type": "paragraph", "w:styleId": "Foo"})
        >>> style.tag.endswith("}style")
        True
    """
    qname = qn(tag)  # validates the prefix before it is used as an nsmap key
    prefix = tag.partition(":")[0]
    nsmap = BUILD_NSMAP if prefix in BUILD_NSMAP else {prefix: NSMAP[prefix]}
    node = etree.Element(qname, nsmap=nsmap)
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
    compiled = _compile_xpath(expr)
    result = compiled(node, **variables)
    if isinstance(result, list):
        return result
    return [result]


@lru_cache(maxsize=512)
def _compile_xpath(expr: str) -> etree.XPath:
    """Compile (and cache) an :class:`lxml.etree.XPath` bound to :data:`NSMAP`.

    The compiled object is reusable across calls with different context
    nodes and ``$name`` variables, so caching on the expression string
    alone is correct — ``NSMAP`` is a module constant. Registry seeding and
    read paths call :func:`xpath` from hot loops with a small set of fixed
    expressions, so the cache turns repeated compilation into a dict hit.
    """
    return etree.XPath(expr, namespaces=NSMAP)


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


# Word's bookmark name rules: letter or underscore first, then letters,
# digits, or underscores; max 40 chars. Worth validating because names
# with spaces or punctuation are rejected by Word's UI but accepted in
# raw OOXML, which turns into a confusing unresolved cross-reference
# rather than an error at authoring time.
_BOOKMARK_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]{0,39}$")


def validate_bookmark_name(name: str, *, arg_name: str = "name") -> str:
    """Check ``name`` against Word's bookmark-name grammar.

    Shared by the three places that accept one: ``bookmarks.add_bookmark``
    (which creates it), ``bookmarks.add_cross_reference`` (which
    references it), and ``publishing.add_caption`` (which brackets a
    caption with one). A name only Word's UI would reject produces a
    silently unresolved field, so all three check up front.

    Args:
        name: The candidate bookmark name.
        arg_name: Caller's parameter name, woven into the error message.

    Returns:
        ``name`` (echoed so the call composes inline).

    Raises:
        ValueError: If ``name`` does not match the grammar.
    """
    if not _BOOKMARK_NAME_RE.match(name):
        raise ValueError(f"{arg_name} {name!r} must match {_BOOKMARK_NAME_RE.pattern}")
    return name


def build_bookmark(
    start_anchor: etree._Element,
    end_anchor: etree._Element,
    *,
    bookmark_id: int,
    name: str,
) -> tuple[etree._Element, etree._Element]:
    """Bracket ``start_anchor``..``end_anchor`` with a bookmark pair.

    Emits ``<w:bookmarkStart>`` immediately before ``start_anchor`` and
    ``<w:bookmarkEnd>`` immediately after ``end_anchor``, both carrying
    ``bookmark_id``. Pass the same element twice to bookmark a single run.

    Lives in ``core`` because two capability packages need it and SPEC
    §9.1 forbids the sibling import: ``bookmarks`` owns the public
    :func:`~docx_plus.bookmarks.add_bookmark`, while ``publishing`` needs
    the same emission to make a caption referenceable — a ``REF`` field
    can only point at a bookmark, never at the caption's ``SEQ`` field.

    Neither the name nor the id is validated here; the caller owns that
    (``bookmarks`` checks the name against the ECMA-376 grammar and
    allocates the id from a registry).

    Args:
        start_anchor: Element the ``bookmarkStart`` is placed before.
        end_anchor: Element the ``bookmarkEnd`` is placed after.
        bookmark_id: Numeric id shared by the pair.
        name: Bookmark name written on the start element.

    Returns:
        The ``(start, end)`` element pair, already inserted.
    """
    bid = str(bookmark_id)
    start = el("w:bookmarkStart", **{"w:id": bid, "w:name": name})
    end = el("w:bookmarkEnd", **{"w:id": bid})

    start_anchor.addprevious(start)
    end_anchor.addnext(end)

    return start, end


def ordered_insert(
    parent: etree._Element,
    child: etree._Element,
    order: tuple[str, ...],
) -> None:
    """Insert ``child`` into ``parent`` at the position dictated by ``order``.

    Removes any existing element in ``parent`` with the same tag first, so
    calling twice with the same tag replaces the previous instance. The
    new element is placed immediately before the first existing sibling
    whose local name comes later in ``order``; a tag absent from ``order``,
    or one with no later sibling present, is appended.

    The stronger sibling of :func:`insert_before_first_anchor`: that one
    takes an explicit list of later siblings and never replaces, this one
    derives the insertion point from a full child-order tuple and is
    idempotent. Reach for this when the parent's complete schema sequence
    is known — ``w:style``, ``w:pPr``, ``w:rPr``, ``w:lvl``.

    Args:
        parent: The element to insert into.
        child: The element to insert. Any existing sibling with the same
            tag is removed first.
        order: Local names in schema order.
    """
    target_local = etree.QName(child.tag).localname
    for existing in parent.findall(child.tag):
        parent.remove(existing)
    try:
        target_idx = order.index(target_local)
    except ValueError:
        parent.append(child)
        return
    for sibling in parent:
        if not isinstance(sibling.tag, str):
            continue
        sibling_local = etree.QName(sibling.tag).localname
        try:
            sibling_idx = order.index(sibling_local)
        except ValueError:
            continue
        if sibling_idx > target_idx:
            sibling.addprevious(child)
            return
    parent.append(child)


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


def body_document_for(proxy: Any, *, operation: str = "this operation") -> Document:
    """Return the main-body :class:`~docx.document.Document` owning ``proxy``.

    python-docx proxies (``Run``, ``Paragraph``, …) inherit ``.part`` from
    ``Parented``. For a proxy in the main document body, ``part`` is the
    ``DocumentPart``, which exposes a ``.document`` property. A header /
    footer proxy is parented to a part that does not, so this raises a
    clear :class:`ValueError` naming the operation that needs the body.

    Shared by the ``comments`` and ``notes`` packages (a ``core`` utility,
    so the duplication SPEC §9.1 would otherwise force is avoided).

    Args:
        proxy: A python-docx proxy exposing ``.part`` (e.g. ``Run``,
            ``Paragraph``).
        operation: Human-readable name of the caller, woven into the error
            message when ``proxy`` is not body-parented.

    Returns:
        The owning :class:`~docx.document.Document`.

    Raises:
        ValueError: If ``proxy`` is not parented to the main document body.
    """
    part: Any = proxy.part
    document = getattr(part, "document", None)
    if document is None:
        raise ValueError(
            f"{operation} only supports the main document body; "
            f"got a proxy parented to {type(part).__name__}"
        )
    return cast("Document", document)


__all__ = [
    "body_document_for",
    "build_bookmark",
    "build_complex_field",
    "el",
    "insert_before_first_anchor",
    "ordered_insert",
    "remove",
    "sub",
    "validate_bookmark_name",
    "xpath",
]
