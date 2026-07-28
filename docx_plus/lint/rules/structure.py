"""Outline and list-structure rules.

Objective defects rather than opinions: an outline that skips a level or a
heading with no text is wrong regardless of house style, so these ship on.

These are cheap only because the cascade resolves ``outline_level`` and
``num_id`` through the style chain — before that, a correctly-styled
``Heading 2`` and a directly-formatted lookalike were indistinguishable.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from docx_plus.lint.models import Issue, Location
from docx_plus.lint.registry import rule

if TYPE_CHECKING:
    from collections.abc import Iterator

    from docx_plus.lint.models import LintContext
    from docx_plus.styles import ResolvedParagraph


# A leading list marker someone typed by hand: "1.", "1)", "(a)", "iv.",
# "- ", "* ", or a bullet glyph. Anchored, and requires trailing space so
# "1.5x faster" is not a list item.
_TYPED_LIST_MARKER = re.compile(
    r"""^\s*(
        \(?\d+[.)]                 # 1.  1)  (1)
      | \(?[a-z][.)]               # a.  a)  (a)
      | \(?(?=[ivx]+[.)])[ivx]+[.)]  # i.  iv)
      | [-*•●▪·]  # - * bullets
    )\s+""",
    re.VERBOSE | re.IGNORECASE,
)

# Word's own heading styles, by resolved id.
_HEADING_ID = re.compile(r"^Heading([1-9])$", re.IGNORECASE)


def _heading_level(resolved: ResolvedParagraph) -> int | None:
    """The 1-based outline level of a heading paragraph, or None.

    Prefers the resolved ``outline_level`` (0-based in OOXML) because it
    walks the style chain, and falls back to the style id for documents
    whose heading styles omit ``w:outlineLvl``.
    """
    level = resolved.formatting.outline_level
    if level is not None and 0 <= level <= 8:
        return level + 1
    match = _HEADING_ID.match(resolved.formatting.style_id or "")
    return int(match.group(1)) if match else None


@rule(
    id="heading-level-skip",
    kind="structural",
    severity="warning",
    description="The outline jumps a level (e.g. Heading 1 straight to Heading 3).",
    tags={"structure", "headings"},
)
def heading_level_skip(ctx: LintContext) -> Iterator[Issue]:
    """Flag a heading more than one level deeper than the one before it.

    A skipped level breaks the document's navigation pane, its table of
    contents, and every accessibility tool that reads the outline.
    """
    previous: int | None = None
    for resolved in ctx.paragraphs:
        level = _heading_level(resolved)
        if level is None:
            continue
        if previous is not None and level > previous + 1:
            yield Issue(
                message=(
                    f"Outline jumps from level {previous} to level {level}, "
                    f"skipping level {previous + 1}."
                ),
                location=Location(
                    paragraph_index=resolved.index,
                    excerpt=ctx.excerpt(resolved.index),
                ),
                observed=f"level {level}",
                expected=f"level {previous + 1}",
            )
        previous = level


@rule(
    id="empty-heading",
    kind="structural",
    severity="warning",
    description="A heading paragraph with no text.",
    tags={"structure", "headings"},
)
def empty_heading(ctx: LintContext) -> Iterator[Issue]:
    """Flag heading paragraphs that carry no text."""
    for resolved in ctx.paragraphs:
        if _heading_level(resolved) is None:
            continue
        if not resolved.text.strip():
            yield Issue(
                message="Heading paragraph has no text.",
                location=Location(paragraph_index=resolved.index),
                observed="empty",
            )


@rule(
    id="manual-list",
    kind="structural",
    severity="warning",
    description="A paragraph starting with a typed list marker instead of real numbering.",
    tags={"structure", "lists"},
)
def manual_list(ctx: LintContext) -> Iterator[Issue]:
    """Flag typed list markers on paragraphs carrying no numbering.

    Typed numbers do not renumber when an item is inserted, do not
    round-trip through a TOC, and are invisible to every tool that
    understands lists.

    Only reachable because ``num_id`` now resolves through the style chain:
    a paragraph styled ``List Bullet`` reports its style's numbering, so a
    genuinely numbered item no longer looks hand-typed.
    """
    for resolved in ctx.paragraphs:
        if resolved.formatting.num_id:
            continue  # genuinely numbered — 0 is the suppression sentinel
        match = _TYPED_LIST_MARKER.match(resolved.text)
        if match is not None:
            yield Issue(
                message="Paragraph begins with a typed list marker but carries no numbering.",
                location=Location(
                    paragraph_index=resolved.index,
                    excerpt=ctx.excerpt(resolved.index),
                ),
                observed=match.group(1),
                expected="a w:numPr list reference",
            )


# `direct-numbering-override` is deliberately absent from this batch. The
# rule is "a paragraph's own w:numPr overrides one its style already
# supplies", and provenance reports only the *winner* — it cannot say what
# the cascade would have produced without the direct layer. Firing on any
# direct numPr instead would flag every list made with `apply_list`, which
# is noise, not a finding. It needs a resolver affordance ("resolve
# beneath the direct layer") that the paragraph-level consistency rules
# share; tracked on the roadmap.
