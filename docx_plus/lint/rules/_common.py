"""Helpers shared by more than one rule module.

Deliberately small. A rule that needs something only it needs should keep
it local; this is for the handful of questions several rules ask in the
same words, where two implementations would be two chances to get it
wrong.
"""

from __future__ import annotations

from collections import Counter
from typing import TYPE_CHECKING, TypeVar

from docx_plus.core.ns import qn

if TYPE_CHECKING:
    from collections.abc import Callable, Hashable, Iterable

    from lxml import etree

    from docx_plus.styles import ResolvedParagraph, ResolvedRun

_T = TypeVar("_T", bound="Hashable")

_BLOCK_TAGS = frozenset({qn("w:p"), qn("w:tbl")})
"""What counts as breaking a run of paragraphs. Anything else between two
``<w:p>`` — a bookmark, a proofing mark, a section property — is not
block-level and leaves them neighbours."""


def paragraph_element(resolved: ResolvedParagraph) -> etree._Element:
    """The underlying ``<w:p>``.

    One place for the ``._p`` reach into python-docx, rather than one per
    callsite.
    """
    return resolved.paragraph._p


def same_container(left: ResolvedParagraph, right: ResolvedParagraph) -> bool:
    """Whether two paragraphs live in the same body or the same table cell.

    Weaker than :func:`document_adjacent`, and the right test for a rule
    tracking a *sequence* rather than a pair — an outline's headings are
    not neighbours, but comparing a body ``Heading 1`` against a
    ``Heading 3`` inside a table cell is still meaningless. Resetting the
    sequence per container is what stops that.
    """
    return paragraph_element(left).getparent() is paragraph_element(right).getparent()


def document_adjacent(previous: ResolvedParagraph, current: ResolvedParagraph) -> bool:
    """Whether two swept paragraphs really are neighbours in the document.

    Consecutive **sweep indices** are not the same thing. The sweep walks
    into table cells, so a body paragraph, a table, and the paragraph after
    it come out as ``... n, n+1, n+2 ...`` with ``n+1`` living inside a
    cell — three positions that share no edge in the rendered document. A
    rule comparing neighbours by index alone therefore reasons across a
    table boundary, and in ``stray-empty-paragraph``'s case planned to
    delete the only paragraph of a ``<w:tc>``, which is not a document Word
    will open without repairing.

    Two paragraphs are adjacent when they share a parent element and
    nothing block-level sits between them. Bookmarks, proofing marks and
    the like are not block-level and do not break a run.
    """
    left = paragraph_element(previous)
    right = paragraph_element(current)
    if left.getparent() is not right.getparent():
        return False

    node = left.getnext()
    while node is not None and node is not right:
        if isinstance(node.tag, str) and node.tag in _BLOCK_TAGS:
            return False
        node = node.getnext()
    return node is right


__all__ = ["document_adjacent", "paragraph_element", "same_container"]


def count_run_values(
    paragraphs: Iterable[ResolvedParagraph],
    key: Callable[[ResolvedRun], _T | None],
) -> Counter[_T]:
    """How often each ``key`` value occurs across the runs that carry text.

    The population behind every "what does this document normally do"
    question: runs of pure whitespace are skipped because their formatting
    is invisible, and a run whose ``key`` is ``None`` is skipped because
    an unset value is not a vote for anything. Callers wanting a subset of
    paragraphs — body text without headings, say — filter before calling.
    """
    return Counter(
        value
        for resolved in paragraphs
        for run in resolved.runs
        if run.run.text.strip() and (value := key(run)) is not None
    )


def dominant(counts: Counter[_T]) -> _T | None:
    """The most common value, or ``None`` when there is nothing to count.

    Ties break by first occurrence, which is document order for every
    counter built by :func:`count_run_values` — the earliest of two
    equally common values wins, so the answer is stable across runs
    rather than an artefact of hashing.
    """
    if not counts:
        return None
    return counts.most_common(1)[0][0]
