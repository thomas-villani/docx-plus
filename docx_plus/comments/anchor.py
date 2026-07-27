"""Anchor and remove comments — the body-side OOXML python-docx skips.

python-docx 1.x writes ``<w:comment>`` into ``comments.xml`` but
omits the three body-side elements that anchor a comment to a text
range — ``w:commentRangeStart``, ``w:commentRangeEnd``, and the
``CommentReference`` marker run. As a result, comments added via
``python-docx`` show up in the review pane but have nothing in the
document text to point at. This module fills that gap.

:func:`add_comment` wraps a run, paragraph, or run-range with the three
body markers and appends a matching ``w:comment`` body to
``comments.xml`` (the comments part is created on first use).

:func:`delete_comment` removes everything :func:`add_comment` wrote.

Every comment written here is also registered as an unresolved thread
root in ``commentsExtended.xml`` — see
:mod:`docx_plus.comments._extended` for why that part exists and
:mod:`docx_plus.comments.threads` for the reply / resolve surface built
on top of it — and given a ``w16cid:durableId`` in ``commentsIds.xml``,
its only identifier stable across edits
(:mod:`docx_plus.comments._ids`).

Author presence (``people.xml``) is *not* written here; it is cosmetic
and needs an identity the library cannot invent. See
:mod:`docx_plus.comments.people`.

This module imports only from ``docx_plus.core`` and the siblings
``docx_plus.comments.registry`` / ``docx_plus.comments._extended`` /
``docx_plus.comments._ids`` (SPEC §9.1).
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass
from typing import TYPE_CHECKING, cast

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import XmlPart
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from lxml import etree

from docx_plus.comments import _extended, _ids
from docx_plus.comments.registry import CommentIdRegistry, DurableIdRegistry
from docx_plus.core import DocxPlusError
from docx_plus.core.ids import ParaIdRegistry
from docx_plus.core.ns import qn
from docx_plus.core.oxml import body_document_for, el, remove, sub, xpath
from docx_plus.core.parts import COMMENTS_SPEC, get_or_create_part

if TYPE_CHECKING:
    from docx.document import Document


CommentTarget = Run | Paragraph | tuple[Run, Run]


class CommentNotFoundError(DocxPlusError, KeyError):
    """Raised when no ``<w:comment>`` with the requested id exists.

    Subclasses :class:`KeyError` so existing ``except KeyError:`` clauses
    still catch it; also :class:`DocxPlusError` per SPEC §9.7.
    """


@dataclass(frozen=True)
class CommentRef:
    """Handle for an inserted comment.

    Attributes:
        comment_id: The ``w:id`` value assigned to the comment.
        body_element: The ``<w:comment>`` lxml element appended to
            ``comments.xml``. Mutate it directly for advanced edits
            (extra paragraphs, formatted runs) the v0.2 surface
            does not yet expose.
    """

    comment_id: int
    body_element: etree._Element


def add_comment(
    target: CommentTarget,
    text: str,
    *,
    author: str = "",
    initials: str | None = None,
    id_registry: CommentIdRegistry | None = None,
    para_id_registry: ParaIdRegistry | None = None,
    durable_id_registry: DurableIdRegistry | None = None,
) -> CommentRef:
    """Anchor a comment to a run, paragraph, or run range.

    Writes the three body-side OOXML elements python-docx skips
    (``w:commentRangeStart``, ``w:commentRangeEnd``, the
    ``CommentReference`` marker run) plus the comment body entry in
    ``comments.xml``. The comments part is created on first use.

    The comment is also registered as an unresolved thread root: its body
    paragraphs are stamped with ``w14:paraId`` and a matching
    ``<w15:commentEx>`` entry is written to ``commentsExtended.xml``
    (created on first use). This is what Word writes for every comment
    it authors, and it is what makes the comment immediately usable with
    :func:`~docx_plus.comments.reply_to_comment` and
    :func:`~docx_plus.comments.resolve_comment`.

    It is also given a ``w16cid:durableId`` in ``commentsIds.xml`` — the
    comment's only identifier stable across edits, since ``w:id`` is a
    position-dependent index Word renumbers and ``w14:paraId`` changes
    whenever the body is rewritten. Read it back through
    :attr:`~docx_plus.comments.AnchoredComment.durable_id`.

    Author presence (``people.xml``) is deliberately *not* written —
    see :func:`~docx_plus.comments.set_author_presence`.

    Args:
        target: Where the comment anchors.

            - A python-docx :class:`~docx.text.run.Run` wraps that one run.
            - A :class:`~docx.text.paragraph.Paragraph` wraps every run
              in the paragraph; the paragraph must contain at least one
              run.
            - A ``(start_run, end_run)`` tuple spans from the start run's
              leading edge to the end run's trailing edge. Both runs
              must already be parented and live in the main document
              body. The caller is responsible for ordering: ``start_run``
              must appear before ``end_run`` in document order. A reversed
              pair is accepted without error but produces a backwards
              range that Word renders as empty.
        text: Comment body text. Whitespace is preserved
            (``xml:space="preserve"``).
        author: Author shown in the review pane. The empty string is
            legal and is what python-docx's own ``add_comment`` uses.
        initials: Author initials shown alongside the author. ``None``
            defaults to the first character of ``author``; pass an
            empty string to suppress the attribute entirely.
        id_registry: Pre-existing registry to share across an editing
            session (useful when inserting many comments). A fresh
            :class:`CommentIdRegistry` is constructed from the target's
            document if not supplied.
        para_id_registry: Pre-existing ``w14:paraId`` allocator, shared
            for the same reason as ``id_registry``. A fresh
            :class:`~docx_plus.core.ids.ParaIdRegistry` is constructed
            from the target's document if not supplied.
        durable_id_registry: Pre-existing ``w16cid:durableId`` allocator,
            shared for the same reason again. A fresh
            :class:`DurableIdRegistry` is constructed from the target's
            document if not supplied.

    Returns:
        A :class:`CommentRef` with the assigned comment id and a handle
        to the comment body element in ``comments.xml``.

    Raises:
        ValueError: If ``target`` is a paragraph with no runs.
        TypeError: If ``target`` is not a Run, Paragraph, or
            ``(Run, Run)`` tuple.

    Example:
        >>> from docx import Document
        >>> from docx_plus.comments import add_comment
        >>> doc = Document()
        >>> p = doc.add_paragraph("Hello world")
        >>> ref = add_comment(p, "Greeting", author="Reviewer")
    """
    start_anchor, end_anchor, doc = _normalize_target(target)

    if id_registry is None:
        id_registry = CommentIdRegistry(doc)
    comment_id = id_registry.next()
    cid = str(comment_id)

    range_start = el("w:commentRangeStart", **{"w:id": cid})
    range_end = el("w:commentRangeEnd", **{"w:id": cid})
    ref_run = _build_reference_run(comment_id)

    start_anchor.addprevious(range_start)
    end_anchor.addnext(range_end)
    range_end.addnext(ref_run)

    _, comments_root = get_or_create_part(doc, COMMENTS_SPEC)
    body = _build_comment_body(comment_id, text, author, initials)
    comments_root.append(body)

    if para_id_registry is None:
        para_id_registry = ParaIdRegistry(doc)
    key = _extended.stamp_para_ids(body, para_id_registry)
    _extended.upsert_comment_ex(doc, key, done=False)
    _ids.upsert_comment_id(doc, key, registry=durable_id_registry)

    return CommentRef(comment_id=comment_id, body_element=body)


def edit_comment(doc: Document, comment_id: int, text: str) -> None:
    """Replace the body text of an existing comment in place.

    Removes all child block-level content of the matching ``<w:comment>``
    element and appends a fresh paragraph with ``text`` as its run. The
    ``<w:comment>`` element's attributes (``w:author``, ``w:date``,
    ``w:initials``) are preserved — only the body content changes. The
    body-side range markers and reference run are also untouched, so the
    comment stays anchored to the same text range, and the comment's
    ``w14:paraId`` is carried onto the rebuilt paragraph so its position
    in the thread and its resolved state survive the edit.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to mutate.
        comment_id: The ``w:id`` of the comment to edit.
        text: New comment body text. Whitespace is preserved
            (``xml:space="preserve"``).

    Raises:
        CommentNotFoundError: If no comment with ``comment_id`` exists,
            including the case where the comments part itself is absent.
            Subclasses :class:`KeyError`, so ``except KeyError`` also
            catches it (SPEC §16).
    """
    cid = str(comment_id)
    try:
        comments_part = cast("XmlPart", doc.part.part_related_by(RT.COMMENTS))
    except KeyError as exc:
        raise CommentNotFoundError(comment_id) from exc

    matches = xpath(comments_part.element, "./w:comment[@w:id=$cid]", cid=cid)
    if not matches:
        raise CommentNotFoundError(comment_id)

    comment_el = matches[0]
    # Capture the thread key before the body is rebuilt. ``commentsExtended.xml``
    # references a comment by the ``w14:paraId`` of its last paragraph, so
    # replacing that paragraph with a freshly built one would orphan the
    # comment's thread entry — replies would detach and a resolved thread
    # would silently reopen. Re-stamping the same value keeps the link.
    key = _extended.thread_key(comment_el)
    # Strip ALL block-level children — ECMA-376 17.13.4.2 (`CT_Comment`)
    # extends `EG_BlockLevelElts`, so a comment authored in Word may
    # legally contain `<w:tbl>`, `<w:sdt>`, `<w:customXml>`, etc. in
    # addition to paragraphs. Filtering to `<w:p>` only would leave
    # those siblings next to the freshly built paragraph. The comment
    # element's own attributes (author / date / initials) live on the
    # element itself, not on its children, so removal is safe.
    for child in list(comment_el):
        remove(child)
    paragraph = _build_comment_paragraph(text)
    if key:
        paragraph.set(qn("w14:paraId"), key)
    comment_el.append(paragraph)


def clear_all_comments(doc: Document, *, remove_part: bool = False) -> None:
    """Remove every comment in the document.

    Single-pass: walks the document body once removing every
    ``<w:commentRangeStart>``, ``<w:commentRangeEnd>``, and
    ``<w:commentReference>`` marker regardless of id, then walks
    ``comments.xml`` once removing every ``<w:comment>`` entry, then
    ``commentsExtended.xml`` and ``commentsIds.xml`` once each removing
    every thread and durable-id entry. Idempotent: a document with no
    comments is a no-op.

    ``people.xml`` is left alone — see
    :func:`~docx_plus.comments.clear_author_presence` for why author
    entries are not pruned automatically.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to scrub.
        remove_part: When ``False`` (default) the now-empty comments,
            commentsExtended, and commentsIds parts are left in place so
            a subsequent :func:`add_comment` reuses them without
            re-creating the relationships. When ``True`` all three parts
            and their relationships are torn down entirely, so the saved
            document carries no comment parts at all — useful when a
            consumer dislikes an empty-but-related part, and the cleaner
            state for a document that is done with comments.
    """
    body = doc.element.body

    for tag_expr in (
        ".//w:commentRangeStart",
        ".//w:commentRangeEnd",
    ):
        for elem in xpath(body, tag_expr):
            remove(elem)

    for ref in xpath(body, ".//w:commentReference"):
        _remove_reference_marker(ref)

    _extended.clear_comments_ex(doc, remove_part=remove_part)
    _ids.clear_comment_ids(doc, remove_part=remove_part)

    try:
        comments_part = cast("XmlPart", doc.part.part_related_by(RT.COMMENTS))
    except KeyError:
        return

    if remove_part:
        _drop_comments_part(doc)
        return

    comments_root = comments_part.element
    for comment_el in list(comments_root.findall(qn("w:comment"))):
        remove(comment_el)


def delete_comment(doc: Document, comment_id: int, *, include_replies: bool = True) -> None:
    """Remove all traces of a comment from the document.

    Removes, for the comment and (by default) every reply beneath it:

    - The ``<w:comment>`` body in ``comments.xml``
    - Every ``<w:commentRangeStart>`` and ``<w:commentRangeEnd>`` marker
      in the document body
    - The reference run hosting ``<w:commentReference>``
    - The ``<w15:commentEx>`` thread entry in ``commentsExtended.xml``

    Idempotent: deleting a comment id that doesn't exist is a no-op.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to mutate.
        comment_id: The ``w:id`` value of the comment to remove.
        include_replies: When ``True`` (default), deleting a comment also
            deletes every reply in its subtree — what Word does when you
            delete a thread root, and what keeps the thread graph from
            retaining entries whose parent is gone. When ``False`` the
            replies survive; each one whose parent was deleted becomes a
            thread root of its own, since a ``paraIdParent`` that no
            comment owns reads as no parent at all.
    """
    targets = [comment_id]
    if include_replies:
        state = _extended.thread_state(doc)
        targets.extend(_extended.descendant_ids(state, comment_id))
    keys, _ = _extended.key_maps(doc)

    for target in targets:
        _delete_one_comment(doc, target, keys.get(target, ""))


def _delete_one_comment(doc: Document, comment_id: int, para_id: str) -> None:
    """Remove a single comment's body, anchors, and thread entry."""
    cid = str(comment_id)
    body = doc.element.body

    for tag_expr in (
        ".//w:commentRangeStart[@w:id=$cid]",
        ".//w:commentRangeEnd[@w:id=$cid]",
    ):
        for elem in xpath(body, tag_expr, cid=cid):
            remove(elem)

    for ref in xpath(body, ".//w:commentReference[@w:id=$cid]", cid=cid):
        _remove_reference_marker(ref)

    _extended.drop_comment_ex(doc, para_id)
    _ids.drop_comment_id(doc, para_id)

    try:
        comments_part = cast("XmlPart", doc.part.part_related_by(RT.COMMENTS))
    except KeyError:
        return
    comments_root = comments_part.element
    for comment_el in xpath(comments_root, "./w:comment[@w:id=$cid]", cid=cid):
        remove(comment_el)


# ---------------------------------------------------------------------------
# Internals.
# ---------------------------------------------------------------------------


def _remove_reference_marker(ref: etree._Element) -> None:
    """Remove a ``<w:commentReference>`` and prune its run only if now empty.

    OOXML permits a single ``<w:r>`` to host the reference marker
    alongside other content — multiple references, or a reference mixed
    with ``<w:t>`` text — and hand-edited or cross-tool round-tripped
    documents do exactly that. Removing the whole parent run
    unconditionally (the pre-fix behaviour) would drop that sibling
    content. So delete only the marker, then remove the run only when
    nothing but an optional ``<w:rPr>`` remains. The reference run
    :func:`add_comment` builds holds only ``<w:rPr>`` + the marker, so
    internally created comments still collapse to nothing as before.
    """
    run = ref.getparent()
    remove(ref)
    if run is None or run.tag != qn("w:r"):
        return
    if all(child.tag == qn("w:rPr") for child in run):
        remove(run)


def _drop_comments_part(doc: Document) -> None:
    """Tear down the comments part and its document-part relationship.

    python-docx serializes only the parts reachable through the
    relationship graph, so dropping the relationship is sufficient to
    keep the comments part out of the saved package. ``drop_rel`` is a
    no-op-safe call: the comments relationship is never referenced by an
    ``r:id`` in ``document.xml`` (its reference count is 0), so it is
    always eligible for removal.
    """
    for rid, rel in list(doc.part.rels.items()):
        if rel.reltype == RT.COMMENTS:
            doc.part.drop_rel(rid)


def _normalize_target(
    target: CommentTarget,
) -> tuple[etree._Element, etree._Element, Document]:
    """Resolve ``target`` to ``(start_anchor, end_anchor, doc)``.

    The caller places ``w:commentRangeStart`` *before* ``start_anchor``
    via :meth:`addprevious` and ``w:commentRangeEnd`` *after*
    ``end_anchor`` via :meth:`addnext`. The returned ``doc`` exposes the
    part graph for the comments-part create-or-reuse step.
    """
    if isinstance(target, Run):
        return target._r, target._r, body_document_for(target, operation="add_comment")

    if isinstance(target, Paragraph):
        runs = [child for child in target._p if child.tag == qn("w:r")]
        if not runs:
            raise ValueError("add_comment requires a paragraph with at least one run")
        return runs[0], runs[-1], body_document_for(target, operation="add_comment")

    if isinstance(target, tuple) and len(target) == 2:
        first, second = target
        if not (isinstance(first, Run) and isinstance(second, Run)):
            raise TypeError("range target must be a tuple of (Run, Run)")
        return first._r, second._r, body_document_for(first, operation="add_comment")

    raise TypeError(
        f"add_comment target must be Run, Paragraph, or (Run, Run); got {type(target).__name__}"
    )


def _build_reference_run(comment_id: int) -> etree._Element:
    """Build the body-side ``<w:r>`` that holds ``<w:commentReference>``."""
    run = el("w:r")
    run_pr = sub(run, "w:rPr")
    sub(run_pr, "w:rStyle", **{"w:val": "CommentReference"})
    sub(run, "w:commentReference", **{"w:id": str(comment_id)})
    return run


def _build_comment_body(
    comment_id: int,
    text: str,
    author: str,
    initials: str | None,
) -> etree._Element:
    """Build the ``<w:comment>`` element appended to ``comments.xml``."""
    attrs = {
        "w:id": str(comment_id),
        "w:author": author,
        "w:date": _now_iso(),
    }
    resolved_initials = author[:1] if initials is None and author else initials
    if resolved_initials:
        attrs["w:initials"] = resolved_initials

    comment = el("w:comment", **attrs)
    comment.append(_build_comment_paragraph(text))
    return comment


def _build_comment_paragraph(text: str) -> etree._Element:
    """Build a comment-body ``<w:p>`` containing ``text`` as a run.

    Shared by :func:`add_comment` (initial insertion) and
    :func:`edit_comment` (in-place replacement). The annotation-ref run
    is included so Word renders the speech-bubble glyph in the comment
    pane.
    """
    p = el("w:p")
    p_pr = sub(p, "w:pPr")
    sub(p_pr, "w:pStyle", **{"w:val": "CommentText"})

    annot_run = sub(p, "w:r")
    annot_pr = sub(annot_run, "w:rPr")
    sub(annot_pr, "w:rStyle", **{"w:val": "CommentReference"})
    sub(annot_run, "w:annotationRef")

    text_run = sub(p, "w:r")
    text_t = sub(text_run, "w:t", **{"xml:space": "preserve"})
    text_t.text = text

    return p


def _now_iso() -> str:
    """``xsd:dateTime`` (UTC, millisecond precision) for ``w:date``.

    Millisecond precision keeps two :func:`add_comment` calls in the same
    wall-clock second from colliding on an identical timestamp. The
    trailing ``Z`` is the canonical UTC designator; ``read_comments``
    parses it back through :func:`datetime.datetime.fromisoformat`.
    """
    now = dt.datetime.now(dt.timezone.utc)
    return now.isoformat(timespec="milliseconds").replace("+00:00", "Z")


__all__ = [
    "CommentNotFoundError",
    "CommentRef",
    "CommentTarget",
    "add_comment",
    "clear_all_comments",
    "delete_comment",
    "edit_comment",
]
