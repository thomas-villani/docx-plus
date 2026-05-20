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

This module imports only from ``docx_plus.core`` and the sibling
``docx_plus.comments.registry`` (SPEC §9.1).
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

from docx_plus.comments.registry import CommentIdRegistry
from docx_plus.core import DocxPlusError
from docx_plus.core.ns import qn
from docx_plus.core.oxml import el, remove, sub, xpath
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
) -> CommentRef:
    """Anchor a comment to a run, paragraph, or run range.

    Writes the three body-side OOXML elements python-docx skips
    (``w:commentRangeStart``, ``w:commentRangeEnd``, the
    ``CommentReference`` marker run) plus the comment body entry in
    ``comments.xml``. The comments part is created on first use.

    Args:
        target: Where the comment anchors.

            - A python-docx :class:`~docx.text.run.Run` wraps that one run.
            - A :class:`~docx.text.paragraph.Paragraph` wraps every run
              in the paragraph; the paragraph must contain at least one
              run.
            - A ``(start_run, end_run)`` tuple spans from the start run's
              leading edge to the end run's trailing edge. Both runs
              must already be parented and live in the main document
              body.
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

    return CommentRef(comment_id=comment_id, body_element=body)


def edit_comment(doc: Document, comment_id: int, text: str) -> None:
    """Replace the body text of an existing comment in place.

    Removes all child paragraphs of the matching ``<w:comment>`` element
    and appends a fresh paragraph with ``text`` as its run. The
    ``<w:comment>`` element's attributes (``w:author``, ``w:date``,
    ``w:initials``) are preserved — only the body content changes. The
    body-side range markers and reference run are also untouched, so the
    comment stays anchored to the same text range.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to mutate.
        comment_id: The ``w:id`` of the comment to edit.
        text: New comment body text. Whitespace is preserved
            (``xml:space="preserve"``).

    Raises:
        CommentNotFoundError: If no comment with ``comment_id`` exists,
            including the case where the comments part itself is absent.
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
    for child in list(comment_el):
        if isinstance(child.tag, str) and etree.QName(child.tag).localname == "p":
            remove(child)
    comment_el.append(_build_comment_paragraph(text))


def clear_all_comments(doc: Document) -> None:
    """Remove every comment in the document.

    Iterates the ``comments.xml`` part for ``w:id`` values and routes
    each through :func:`delete_comment`. The comments part itself is
    left in place (empty) so subsequent calls to :func:`add_comment`
    reuse it without re-creating the relationship. Idempotent: a
    document with no comments is a no-op.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to scrub.
    """
    try:
        comments_part = cast("XmlPart", doc.part.part_related_by(RT.COMMENTS))
    except KeyError:
        return
    comments_root = comments_part.element
    for comment_el in list(comments_root.findall(qn("w:comment"))):
        raw_id = comment_el.get(qn("w:id"))
        if raw_id is None:
            continue
        try:
            comment_id = int(raw_id)
        except ValueError:
            continue
        delete_comment(doc, comment_id)


def delete_comment(doc: Document, comment_id: int) -> None:
    """Remove all traces of a comment from the document.

    Removes:

    - The ``<w:comment>`` body in ``comments.xml``
    - Every ``<w:commentRangeStart>`` and ``<w:commentRangeEnd>`` marker
      in the document body
    - The reference run hosting ``<w:commentReference>``

    Idempotent: deleting a comment id that doesn't exist is a no-op.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to mutate.
        comment_id: The ``w:id`` value of the comment to remove.
    """
    cid = str(comment_id)
    body = doc.element.body

    for tag_expr in (
        ".//w:commentRangeStart[@w:id=$cid]",
        ".//w:commentRangeEnd[@w:id=$cid]",
    ):
        for elem in xpath(body, tag_expr, cid=cid):
            remove(elem)

    for ref in xpath(body, ".//w:commentReference[@w:id=$cid]", cid=cid):
        run = ref.getparent()
        if run is not None and run.tag == qn("w:r"):
            remove(run)
        else:
            remove(ref)

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
        return target._r, target._r, _doc_for(target)

    if isinstance(target, Paragraph):
        runs = [child for child in target._p if child.tag == qn("w:r")]
        if not runs:
            raise ValueError("add_comment requires a paragraph with at least one run")
        return runs[0], runs[-1], _doc_for(target)

    if isinstance(target, tuple) and len(target) == 2:
        first, second = target
        if not (isinstance(first, Run) and isinstance(second, Run)):
            raise TypeError("range target must be a tuple of (Run, Run)")
        return first._r, second._r, _doc_for(first)

    raise TypeError(
        f"add_comment target must be Run, Paragraph, or (Run, Run); got {type(target).__name__}"
    )


def _doc_for(proxy: Run | Paragraph) -> Document:
    """Return the :class:`Document` containing ``proxy``.

    python-docx proxies inherit ``.part`` from :class:`Parented`. For a
    proxy in the main body, ``part`` is the :class:`DocumentPart` which
    exposes a ``.document`` property; for headers/footers it would not,
    which is why we bail with a clear error.
    """
    part = proxy.part
    document = getattr(part, "document", None)
    if document is None:
        raise ValueError(
            "add_comment only supports the main document body in v0.2; "
            f"got proxy parented to {type(part).__name__}"
        )
    return cast("Document", document)


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
    """``xsd:dateTime`` (UTC) for the ``w:date`` attribute."""
    return dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


__all__ = [
    "CommentNotFoundError",
    "CommentRef",
    "CommentTarget",
    "add_comment",
    "clear_all_comments",
    "delete_comment",
    "edit_comment",
]
