"""Threaded comments — replies, resolve / reopen, and thread reads.

Word has modelled comments as *threads* since 2013: a root comment plus
ordered replies, with a resolved flag driving the review pane's
"Resolve" button. python-docx exposes none of that, and neither did
``docx_plus`` before v0.4 — :mod:`docx_plus.comments.anchor` writes a
flat list of comments.

This module adds the missing three operations:

- :func:`reply_to_comment` — attach a reply to an existing comment
- :func:`resolve_comment` / :func:`reopen_comment` — toggle a thread's
  resolved state
- :func:`read_threads` — read comments back as nested threads

The thread graph itself lives in ``commentsExtended.xml``; see
:mod:`docx_plus.comments._extended` for that plumbing. This module
imports only from ``docx_plus.core`` and its ``docx_plus.comments``
siblings (SPEC §9.1).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from lxml import etree

from docx_plus.comments import _extended
from docx_plus.comments.anchor import (
    CommentNotFoundError,
    CommentRef,
    _build_comment_body,
    _build_reference_run,
)
from docx_plus.comments.read import AnchoredComment, read_comments
from docx_plus.comments.registry import CommentIdRegistry
from docx_plus.core.ids import ParaIdRegistry
from docx_plus.core.oxml import el, xpath

if TYPE_CHECKING:
    from docx.document import Document


@dataclass(frozen=True)
class CommentThread:
    """A root comment and every reply beneath it.

    Attributes:
        root: The thread's root comment.
        replies: Every reply in the thread, breadth-first from the root.
            Word only nests one level deep, but ``commentsExtended.xml``
            permits a chain, so a deeper graph is flattened here rather
            than dropped.
        resolved: Whether the thread is marked resolved (``w15:done``).
            Word resolves whole threads, so this reflects the root's
            flag.
    """

    root: AnchoredComment
    replies: tuple[AnchoredComment, ...]
    resolved: bool


def reply_to_comment(
    doc: Document,
    parent_id: int,
    text: str,
    *,
    author: str = "",
    initials: str | None = None,
    id_registry: CommentIdRegistry | None = None,
    para_id_registry: ParaIdRegistry | None = None,
) -> CommentRef:
    """Add a reply beneath an existing comment.

    Writes a new comment body, links it to ``parent_id`` in
    ``commentsExtended.xml``, and mirrors the parent's body-side anchors
    so the reply spans the same text range — which is how Word renders a
    thread as one balloon rather than two.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to mutate.
        parent_id: ``w:id`` of the comment being replied to. Replying to
            a reply is allowed; the new comment is parented to the
            comment you name, not silently re-pointed at the root.
        text: Reply body text. Whitespace is preserved
            (``xml:space="preserve"``).
        author: Author shown in the review pane.
        initials: Author initials. ``None`` defaults to the first
            character of ``author``; pass an empty string to suppress
            the attribute entirely.
        id_registry: Pre-existing comment-id registry to share across an
            editing session.
        para_id_registry: Pre-existing ``w14:paraId`` allocator to share
            across an editing session.

    Returns:
        A :class:`~docx_plus.comments.CommentRef` for the new reply.

    Raises:
        CommentNotFoundError: If no comment with ``parent_id`` exists,
            including the case where the comments part itself is absent.
            Subclasses :class:`KeyError`.

    Note:
        If the parent is *orphaned* — present in ``comments.xml`` but
        with no body-side range markers, the state
        :func:`~docx_plus.comments.read_comments` reports with
        ``paragraph_index=-1`` — there is no range for the reply to
        mirror, so the reply is written orphaned too. It is a valid
        thread member but, like its parent, invisible in Word until the
        anchors are repaired.

    Example:
        >>> from docx import Document
        >>> from docx_plus.comments import add_comment, reply_to_comment
        >>> doc = Document()
        >>> p = doc.add_paragraph("Hello world")
        >>> root = add_comment(p, "Is this right?", author="Reviewer")
        >>> reply = reply_to_comment(doc, root.comment_id, "Yes.", author="Author")
    """
    parents = _extended.comment_elements(doc)
    parent_el = parents.get(parent_id)
    if parent_el is None:
        raise CommentNotFoundError(parent_id)

    if id_registry is None:
        id_registry = CommentIdRegistry(doc)
    comment_id = id_registry.next()

    comments_root = parent_el.getparent()
    if comments_root is None:  # pragma: no cover - a parsed comment is always parented
        raise CommentNotFoundError(parent_id)
    body = _build_comment_body(comment_id, text, author, initials)
    comments_root.append(body)

    if para_id_registry is None:
        para_id_registry = ParaIdRegistry(doc)
    parent_key = _extended.stamp_para_ids(parent_el, para_id_registry)
    reply_key = _extended.stamp_para_ids(body, para_id_registry)

    # The parent may predate this library (python-docx, or Word before
    # 2013) and so have had no entry until the stamp above. Materialize
    # it as an unresolved root before pointing the reply at it, or the
    # reply would reference a key with no entry.
    _extended.upsert_comment_ex(doc, parent_key)
    _extended.upsert_comment_ex(doc, reply_key, parent_para_id=parent_key, done=False)

    _mirror_anchors(doc, parent_id, comment_id)

    return CommentRef(comment_id=comment_id, body_element=body)


def resolve_comment(doc: Document, comment_id: int) -> None:
    """Mark the thread containing ``comment_id`` as resolved.

    Resolution is a property of the *thread*, not of one comment — Word's
    Resolve button greys out the root and every reply together — so this
    sets ``w15:done="1"`` across the whole thread no matter which member
    you name.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to mutate.
        comment_id: Any comment in the thread to resolve.

    Raises:
        CommentNotFoundError: If no comment with ``comment_id`` exists.
    """
    _set_thread_done(doc, comment_id, done=True)


def reopen_comment(doc: Document, comment_id: int) -> None:
    """Mark the thread containing ``comment_id`` as unresolved.

    The exact inverse of :func:`resolve_comment`, with the same
    thread-wide semantics.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to mutate.
        comment_id: Any comment in the thread to reopen.

    Raises:
        CommentNotFoundError: If no comment with ``comment_id`` exists.
    """
    _set_thread_done(doc, comment_id, done=False)


def read_threads(doc: Document) -> list[CommentThread]:
    """Return every comment in ``doc`` grouped into threads.

    A document with no ``commentsExtended.xml`` — anything written by
    python-docx, or by Word before 2013 — yields one single-comment
    thread per comment, all unresolved. That is the correct reading:
    without the extended part there is no threading information to
    report.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to scan.

    Returns:
        One :class:`CommentThread` per root comment, in ``comments.xml``
        order. Returns ``[]`` if the document has no comments part.
    """
    comments = read_comments(doc)
    if not comments:
        return []

    by_id = {comment.comment_id: comment for comment in comments}
    state = _extended.thread_state(doc)

    threads: list[CommentThread] = []
    for comment in comments:
        if state.get(comment.comment_id, (None, False))[0] is not None:
            continue  # a reply — it is emitted under its root
        replies = tuple(
            by_id[reply_id]
            for reply_id in _extended.descendant_ids(state, comment.comment_id)
            if reply_id in by_id
        )
        threads.append(CommentThread(root=comment, replies=replies, resolved=comment.resolved))
    return threads


# ---------------------------------------------------------------------------
# Internals.
# ---------------------------------------------------------------------------


def _set_thread_done(doc: Document, comment_id: int, *, done: bool) -> None:
    """Set ``w15:done`` on every comment in ``comment_id``'s thread."""
    elements = _extended.comment_elements(doc)
    if comment_id not in elements:
        raise CommentNotFoundError(comment_id)

    # Stamp first: a comment that has never been through this library has
    # no paraId, hence no entry to flip.
    registry = ParaIdRegistry(doc)
    for element in elements.values():
        _extended.stamp_para_ids(element, registry)

    state = _extended.thread_state(doc)
    root_id = _extended.root_id_of(state, comment_id)
    keys, _ = _extended.key_maps(doc)

    for member_id in [root_id, *_extended.descendant_ids(state, root_id)]:
        _extended.upsert_comment_ex(doc, keys.get(member_id, ""), done=done)


def _mirror_anchors(doc: Document, parent_id: int, reply_id: int) -> None:
    """Give ``reply_id`` the same body-side range as ``parent_id``.

    Word nests the markers rather than duplicating the span side by side:
    the reply's ``commentRangeStart`` sits just inside the parent's, and
    its ``commentRangeEnd`` plus reference run follow the parent's
    reference run. Written out, a two-comment thread over "Hello" is::

        <w:commentRangeStart w:id="1"/>
        <w:commentRangeStart w:id="2"/>
        <w:r><w:t>Hello</w:t></w:r>
        <w:commentRangeEnd w:id="1"/>
        <w:r><w:commentReference w:id="1"/></w:r>
        <w:commentRangeEnd w:id="2"/>
        <w:r><w:commentReference w:id="2"/></w:r>

    A parent with no anchors (an orphaned comment) leaves the reply
    orphaned too — documented on :func:`reply_to_comment`.
    """
    body = doc.element.body
    parent_cid = str(parent_id)
    reply_cid = str(reply_id)

    starts = xpath(body, ".//w:commentRangeStart[@w:id=$cid]", cid=parent_cid)
    ends = xpath(body, ".//w:commentRangeEnd[@w:id=$cid]", cid=parent_cid)
    if not starts or not ends:
        return

    starts[0].addnext(el("w:commentRangeStart", **{"w:id": reply_cid}))

    reply_end = el("w:commentRangeEnd", **{"w:id": reply_cid})
    _last_marker_after(body, ends[0], parent_cid).addnext(reply_end)
    reply_end.addnext(_build_reference_run(reply_id))


def _last_marker_after(
    body: etree._Element,
    parent_end: etree._Element,
    parent_cid: str,
) -> etree._Element:
    """Return the node the reply's ``commentRangeEnd`` should follow.

    That is the parent's reference run when one exists — keeping the
    parent's end/reference pair adjacent, the way Word writes it — and
    the parent's ``commentRangeEnd`` otherwise (a range whose reference
    marker was stripped by another tool).
    """
    refs = xpath(body, ".//w:commentReference[@w:id=$cid]", cid=parent_cid)
    if not refs:
        return parent_end
    run = refs[0].getparent()
    return run if run is not None else parent_end


__all__ = [
    "CommentThread",
    "read_threads",
    "reopen_comment",
    "reply_to_comment",
    "resolve_comment",
]
