"""Bookmark anchoring — body-side ``<w:bookmarkStart>`` / ``<w:bookmarkEnd>``.

A bookmark is a pair of empty marker elements with a shared ``w:id`` and
a human-readable ``w:name``. python-docx provides no abstraction for
either, so this module fills the gap with :func:`add_bookmark` and
:func:`delete_bookmark`. Cross-references key off the bookmark *name*,
which is what ``REF`` / ``PAGEREF`` field instructions accept.

This module imports only from ``docx_plus.core`` and the sibling
``docx_plus.bookmarks.registry`` (SPEC §9.1).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING, cast

from docx.text.paragraph import Paragraph
from docx.text.run import Run
from lxml import etree

from docx_plus.bookmarks.registry import BookmarkIdRegistry
from docx_plus.core.ns import qn
from docx_plus.core.oxml import el, remove, xpath

if TYPE_CHECKING:
    from docx.document import Document


BookmarkTarget = Run | Paragraph | tuple[Run, Run]

# Word's bookmark name rules: letter or underscore first, then letters,
# digits, or underscores; max 40 chars. We validate to keep cross-refs
# resolvable — names with spaces or punctuation are silently rejected by
# Word's UI but accepted in raw OOXML, which leads to confusing failures.
_BOOKMARK_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]{0,39}$")


@dataclass(frozen=True)
class BookmarkRef:
    """Handle for an inserted bookmark.

    Attributes:
        bookmark_id: The ``w:id`` value shared by the start and end
            markers.
        name: The ``w:name`` attribute. Cross-references key off the
            name, not the id.
        start_element: The ``<w:bookmarkStart>`` lxml element.
        end_element: The ``<w:bookmarkEnd>`` lxml element.
    """

    bookmark_id: int
    name: str
    start_element: etree._Element
    end_element: etree._Element


def add_bookmark(
    target: BookmarkTarget,
    name: str,
    *,
    id_registry: BookmarkIdRegistry | None = None,
) -> BookmarkRef:
    """Anchor a bookmark to a run, paragraph, or run range.

    Writes a paired ``<w:bookmarkStart>`` / ``<w:bookmarkEnd>``
    bracketing the target. The bookmark id is minted from
    ``id_registry`` (or a fresh one if not supplied). The name is
    validated against Word's bookmark name rules: it must start with a
    letter or underscore, contain only letters, digits, and underscores,
    and be at most 40 characters long.

    Args:
        target: Where the bookmark anchors. Same shapes as
            :func:`docx_plus.comments.add_comment` — a single ``Run``, a
            ``Paragraph`` (must have at least one run), or a
            ``(start_run, end_run)`` tuple.
        name: Bookmark name. Must match
            ``[A-Za-z_][A-Za-z0-9_]{0,39}``. Names violating Word's
            rules silently break cross-references, so this is enforced.
        id_registry: Pre-existing registry to share across an editing
            session.

    Returns:
        A :class:`BookmarkRef` capturing the assigned id and the body
        elements.

    Raises:
        ValueError: For invalid names, empty paragraph targets, or
            unsupported target shapes.

    Example:
        >>> from docx import Document
        >>> from docx_plus.bookmarks import add_bookmark
        >>> doc = Document()
        >>> p = doc.add_paragraph("Section 1 intro")
        >>> ref = add_bookmark(p, "section_1_intro")
    """
    if not _BOOKMARK_NAME_RE.match(name):
        raise ValueError(
            f"bookmark name {name!r} must match {_BOOKMARK_NAME_RE.pattern}"
        )

    start_anchor, end_anchor, doc = _normalize_target(target)

    if id_registry is None:
        id_registry = BookmarkIdRegistry(doc)
    bookmark_id = id_registry.next()
    bid = str(bookmark_id)

    start = el("w:bookmarkStart", **{"w:id": bid, "w:name": name})
    end = el("w:bookmarkEnd", **{"w:id": bid})

    start_anchor.addprevious(start)
    end_anchor.addnext(end)

    return BookmarkRef(
        bookmark_id=bookmark_id,
        name=name,
        start_element=start,
        end_element=end,
    )


def delete_bookmark(doc: Document, name: str) -> None:
    """Remove every bookmark with the given name from ``doc``.

    Idempotent — removing a missing bookmark is a no-op. Removing by
    *name* (not id) matches the cross-reference key, so a stale name
    referenced by a ``REF`` field can be cleared with a single call.

    Args:
        doc: A python-docx Document.
        name: Bookmark name to remove. Matches case-sensitively.
    """
    body = doc.element.body
    starts = xpath(body, ".//w:bookmarkStart[@w:name=$name]", name=name)
    ids = {s.get(qn("w:id")) for s in starts}
    for start in starts:
        remove(start)
    if not ids:
        return
    # Match each end by id (bookmarkEnd has no name attribute).
    for end in xpath(body, ".//w:bookmarkEnd"):
        if end.get(qn("w:id")) in ids:
            remove(end)


def _normalize_target(
    target: BookmarkTarget,
) -> tuple[etree._Element, etree._Element, Document]:
    """Resolve ``target`` to ``(start_anchor, end_anchor, doc)``.

    Mirrors :func:`docx_plus.comments.anchor._normalize_target`. The
    returned ``doc`` is the python-docx Document containing the
    target — used to construct an :class:`BookmarkIdRegistry` if one
    wasn't supplied.
    """
    if isinstance(target, Run):
        return target._r, target._r, _doc_for(target)

    if isinstance(target, Paragraph):
        runs = [child for child in target._p if child.tag == qn("w:r")]
        if not runs:
            raise ValueError("add_bookmark requires a paragraph with at least one run")
        return runs[0], runs[-1], _doc_for(target)

    if isinstance(target, tuple) and len(target) == 2:
        first, second = target
        if not (isinstance(first, Run) and isinstance(second, Run)):
            raise TypeError("range target must be a tuple of (Run, Run)")
        return first._r, second._r, _doc_for(first)

    raise TypeError(
        f"add_bookmark target must be Run, Paragraph, or (Run, Run); "
        f"got {type(target).__name__}"
    )


def _doc_for(proxy: Run | Paragraph) -> Document:
    """Return the :class:`Document` containing ``proxy``.

    Same shape as ``docx_plus.comments.anchor._doc_for``.
    """
    part = proxy.part
    document = getattr(part, "document", None)
    if document is None:
        raise ValueError(
            "add_bookmark only supports the main document body in v0.2; "
            f"got proxy parented to {type(part).__name__}"
        )
    return cast("Document", document)


__all__ = [
    "BookmarkRef",
    "BookmarkTarget",
    "add_bookmark",
    "delete_bookmark",
]
