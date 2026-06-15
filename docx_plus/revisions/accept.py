"""Accept or reject tracked changes — resolve revision marks into final text.

Accepting a revision keeps the edit it records; rejecting restores the
pre-edit state. The transform differs per element type:

================  ==========================  ===========================
Element           Accept                      Reject
================  ==========================  ===========================
``w:ins``         unwrap (keep runs)          remove (drop element+runs)
``w:del``         remove                      unwrap + ``w:delText``→``w:t``
``w:moveFrom``    remove                      unwrap (+ retag)
``w:moveTo``      unwrap                      remove
move markers      remove                      remove
``w:rPrChange``   remove marker               restore recorded old ``w:rPr``
``w:pPrChange``   remove marker               restore recorded old ``w:pPr``
para-mark ins/del remove mark (safe)          remove mark (safe)
================  ==========================  ===========================

Run-level insertions and deletions are handled fully. Move and
property-change marks get the safe, non-structural transform above. The
one genuinely structural case — accepting a *paragraph-mark* deletion
should merge two paragraphs — ships a non-corrupting fallback (the mark is
removed, the text is left intact) rather than attempting the merge; true
merge/split is deferred (ROADMAP v0.3+).

This module imports only from ``docx_plus.core`` and the sibling
``docx_plus.revisions.mark`` (SPEC §9.1).
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from docx_plus.core.ns import qn
from docx_plus.core.oxml import el, remove, xpath
from docx_plus.revisions.mark import RevisionNotFoundError

if TYPE_CHECKING:
    from docx.document import Document


# Every revision element type, by ``prefix:local`` tag.
_REVISION_TAGS: tuple[str, ...] = (
    "w:ins",
    "w:del",
    "w:moveFrom",
    "w:moveTo",
    "w:moveFromRangeStart",
    "w:moveFromRangeEnd",
    "w:moveToRangeStart",
    "w:moveToRangeEnd",
    "w:rPrChange",
    "w:pPrChange",
)


def accept_revision(doc: Document, revision_id: int) -> None:
    """Accept the revision(s) with ``revision_id``, keeping the recorded edit.

    Locates every body element carrying ``@w:id == revision_id`` and applies
    the per-type accept transform. To fully resolve a *move* (whose wrapper
    and range markers carry distinct ids) prefer
    :func:`accept_all_revisions`.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to mutate.
        revision_id: The ``w:id`` of the revision to accept.

    Raises:
        RevisionNotFoundError: If no revision element carries ``revision_id``.
            Subclasses :class:`KeyError`.
    """
    _resolve_by_id(doc, revision_id, accept=True)


def reject_revision(doc: Document, revision_id: int) -> None:
    """Reject the revision(s) with ``revision_id``, restoring the prior state.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to mutate.
        revision_id: The ``w:id`` of the revision to reject.

    Raises:
        RevisionNotFoundError: If no revision element carries ``revision_id``.
    """
    _resolve_by_id(doc, revision_id, accept=False)


def accept_all_revisions(doc: Document) -> None:
    """Accept every tracked change in ``doc``.

    Idempotent: a document with no revisions is left unchanged. Revisions
    are processed innermost-first so a nested revision is resolved before
    its container.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to mutate.
    """
    _resolve_all(doc, accept=True)


def reject_all_revisions(doc: Document) -> None:
    """Reject every tracked change in ``doc``, restoring the pre-edit text.

    Idempotent. Revisions are processed innermost-first.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to mutate.
    """
    _resolve_all(doc, accept=False)


# ---------------------------------------------------------------------------
# Drivers.
# ---------------------------------------------------------------------------


def _resolve_by_id(doc: Document, revision_id: int, *, accept: bool) -> None:
    body = doc.element.body
    matches: list[etree._Element] = []
    for tag in _REVISION_TAGS:
        matches.extend(xpath(body, f".//{tag}[@w:id=$rid]", rid=str(revision_id)))
    if not matches:
        raise RevisionNotFoundError(revision_id)
    for elem in sorted(matches, key=_depth, reverse=True):
        _apply(elem, accept=accept)


def _resolve_all(doc: Document, *, accept: bool) -> None:
    body = doc.element.body
    elements: list[etree._Element] = []
    for tag in _REVISION_TAGS:
        elements.extend(xpath(body, f".//{tag}"))
    for elem in sorted(elements, key=_depth, reverse=True):
        _apply(elem, accept=accept)


# ---------------------------------------------------------------------------
# Per-element transform.
# ---------------------------------------------------------------------------


def _apply(elem: etree._Element, *, accept: bool) -> None:
    """Dispatch ``elem`` to its accept/reject transform by tag."""
    tag = etree.QName(elem).localname

    if tag in ("ins", "del") and _is_property_mark(elem):
        # Paragraph/run-mark revision — safe fallback: drop the mark without
        # attempting a structural paragraph merge/split (deferred).
        remove(elem)
        return

    if tag == "ins":
        _unwrap(elem) if accept else remove(elem)
    elif tag == "del":
        if accept:
            remove(elem)
        else:
            _retag_runs(_unwrap(elem), "w:delText", "w:t")
    elif tag == "moveFrom":
        if accept:
            remove(elem)
        else:
            _retag_runs(_unwrap(elem), "w:delText", "w:t")
    elif tag == "moveTo":
        _unwrap(elem) if accept else remove(elem)
    elif tag in ("moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd"):
        # Range delimiters carry no content; resolving the move removes them.
        remove(elem)
    elif tag == "rPrChange":
        _resolve_property_change(elem, "w:rPr", accept=accept)
    elif tag == "pPrChange":
        _resolve_property_change(elem, "w:pPr", accept=accept)


def _resolve_property_change(marker: etree._Element, container_tag: str, *, accept: bool) -> None:
    """Accept/reject a ``w:rPrChange`` / ``w:pPrChange`` marker.

    Accept keeps the current (parent) properties and drops the marker.
    Reject replaces the parent property container's children with the
    recorded *old* properties stored inside the marker — a whole-block swap
    that needs no schema-order reasoning because Word wrote the recorded
    block in order.
    """
    if accept:
        remove(marker)
        return

    outer = marker.getparent()  # the live w:rPr / w:pPr
    if outer is None:
        return
    old = marker.find(qn(container_tag))  # recorded prior props
    for child in list(outer):
        remove(child)
    if old is not None:
        for child in list(old):
            outer.append(child)


# ---------------------------------------------------------------------------
# lxml primitives.
# ---------------------------------------------------------------------------


def _unwrap(wrapper: etree._Element) -> list[etree._Element]:
    """Replace ``wrapper`` with its children, preserving order; return them.

    Returns the lifted child elements (now parented where the wrapper was)
    so callers can scope follow-up work — e.g. retagging ``w:delText`` —
    to exactly the unwrapped content and not to unrelated siblings.
    """
    parent = wrapper.getparent()
    if parent is None:
        return []
    children = list(wrapper)
    for child in children:
        wrapper.addprevious(child)  # moves an already-parented node before wrapper
    parent.remove(wrapper)
    return children


def _retag_runs(runs: list[etree._Element], from_tag: str, to_tag: str) -> None:
    """Retag every ``from_tag`` within each element of ``runs`` to ``to_tag``.

    Scoped to the given subtrees (the just-unwrapped runs) so a sibling
    deletion's ``w:delText`` in the same paragraph is left untouched.
    """
    for run in runs:
        for old in xpath(run, f".//{from_tag}"):
            new = el(to_tag)
            space = old.get(qn("xml:space"))
            if space is not None:
                new.set(qn("xml:space"), space)
            new.text = old.text
            old.addprevious(new)
            remove(old)


def _is_property_mark(elem: etree._Element) -> bool:
    """True if a ``w:ins`` / ``w:del`` marks a paragraph/run mark, not content."""
    parent = elem.getparent()
    if parent is None:
        return False
    return parent.tag in (qn("w:rPr"), qn("w:pPr"))


def _depth(elem: etree._Element) -> int:
    """Number of ancestors of ``elem`` (used to process innermost marks first)."""
    depth = 0
    node = elem.getparent()
    while node is not None:
        depth += 1
        node = node.getparent()
    return depth


__all__ = [
    "accept_all_revisions",
    "accept_revision",
    "reject_all_revisions",
    "reject_revision",
]
