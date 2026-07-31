"""Outline and list-structure rules.

Objective defects rather than opinions: an outline that skips a level or a
heading with no text is wrong regardless of house style, so these ship on.

These are cheap only because the cascade resolves ``outline_level`` and
``num_id`` through the style chain — before that, a correctly-styled
``Heading 2`` and a directly-formatted lookalike were indistinguishable.
"""

from __future__ import annotations

import re
from collections import Counter
from typing import TYPE_CHECKING

from docx_plus.lint.models import Fix, FixOperation, Issue, Location
from docx_plus.lint.registry import rule
from docx_plus.styles import list_styles

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

# Longer than this and it reads as a sentence, not a heading. Generous
# enough for a real one; short enough to exclude most prose.
_HEADING_LIKE_MAX_CHARS = 80


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

    Report-only: promoting this heading and demoting the one above it are
    both valid repairs and they produce different documents.
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
    """Flag heading paragraphs that carry no text.

    Report-only: an empty heading is either a stray paragraph to delete or
    a section whose title never got typed, and the document cannot say
    which.
    """
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

    Report-only: the repair strips the typed marker and applies a real
    list, and which list — an existing definition, a new one, at which
    level — is a choice the paragraph does not contain.
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


@rule(
    id="direct-numbering-override",
    kind="consistency",
    severity="warning",
    description="A paragraph's own numbering overrides the list its style supplies.",
    tags={"structure", "lists", "styles"},
)
def direct_numbering_override(ctx: LintContext) -> Iterator[Issue]:
    """Flag a direct ``w:numPr`` fighting one the paragraph's style supplies.

    A ``List Bullet`` paragraph pointed at some other ``numId`` renders as
    a list either way, so nothing looks wrong — until the style's list is
    restructured and these paragraphs do not follow, or a second "1." list
    appears where one was intended.

    Needs a resolve the sweep does not precompute: the paragraph's baseline
    stops below ``directParagraph``, which is *above* numbering, so it
    still carries the direct reference. ``stop_below="numbering"`` is the
    one that reports what the style alone would give. That is why the
    resolver splits the numbering layer in two — and it is only asked for
    the paragraphs that actually carry a direct reference, since it costs a
    full cascade walk each.
    """
    for resolved in ctx.paragraphs:
        provenance = resolved.formatting.provenance or {}
        source = provenance.get("num_id")
        if source is None or source.layer != "numbering":
            continue  # no numbering, or it came from the style chain

        from_style = ctx.resolve(resolved.paragraph, stop_below="numbering").num_id
        if from_style is None or from_style == resolved.formatting.num_id:
            continue

        direct = resolved.formatting.num_id
        suppressed = direct == 0
        yield Issue(
            message=(
                f"Paragraph suppresses the numbering its style supplies (list {from_style})."
                if suppressed
                else (
                    f"Paragraph points at list {direct} directly, overriding "
                    f"list {from_style} from its style."
                )
            ),
            location=Location(
                paragraph_index=resolved.index,
                style_id=resolved.formatting.style_id,
                excerpt=ctx.excerpt(resolved.index),
            ),
            observed=f"numId={direct}",
            expected=f"numId={from_style} (from the style)",
            # numId 0 is the ECMA-376 17.9.18 opt-out sentinel and the one
            # legitimate reason to override: it is the only way to take a
            # single paragraph out of a style's list.
            severity="info" if suppressed else None,
            # ...which is also why the suppressed case carries no fix. The
            # repair would put the paragraph back in the list, undoing
            # something someone did deliberately. Reporting it is enough.
            fix=None
            if suppressed
            else Fix(
                summary=f"Drop the direct numbering so the style's list {from_style} applies.",
                safety="review",
                operations=(
                    FixOperation(
                        op="clear-paragraph-numbering",
                        args={"paragraph_index": resolved.index},
                    ),
                ),
            ),
        )


@rule(
    id="list-numbering-continuity",
    kind="structural",
    severity="warning",
    description="Consecutive list items at one level belong to different lists.",
    tags={"structure", "lists"},
)
def list_numbering_continuity(ctx: LintContext) -> Iterator[Issue]:
    """Flag an unbroken run of list items split across several ``numId``s.

    The "three separate 1. lists" footgun. Each ``numId`` is an independent
    list with its own counter, so items that look like one list restart
    numbering partway down. It is invisible on a bulleted list and obvious
    — once printed — on a numbered one.

    Only adjacent items at the **same level** are compared. A sublist is a
    different level and legitimately a different list, and a run broken by
    body text is a deliberate restart often enough not to call it.

    Reachable only because ``num_id`` resolves through the style chain: a
    correctly-styled ``List Number`` paragraph reports the list its style
    supplies rather than nothing at all.

    Report-only. Pointing the run of items at one list is the repair, but
    the two lists have their own definitions — glyphs, indents, start
    values — so whichever survives changes how the other half looks. And
    where the numbering comes from a style, the honest repair is to the
    style rather than to these paragraphs.
    """
    previous: ResolvedParagraph | None = None
    for resolved in ctx.paragraphs:
        num_id = resolved.formatting.num_id
        if not num_id:  # None, or the 0 opt-out sentinel
            previous = None
            continue

        if (
            previous is not None
            and previous.formatting.num_level == resolved.formatting.num_level
            and previous.formatting.num_id != num_id
        ):
            yield Issue(
                message=(
                    f"List item belongs to list {num_id}, but the item directly "
                    f"above it is in list {previous.formatting.num_id}; the two "
                    f"number independently."
                ),
                location=Location(
                    paragraph_index=resolved.index,
                    style_id=resolved.formatting.style_id,
                    excerpt=ctx.excerpt(resolved.index),
                ),
                observed=f"numId={num_id}",
                expected=f"numId={previous.formatting.num_id}",
            )
        previous = resolved


@rule(
    id="manual-heading-formatting",
    kind="structural",
    severity="warning",
    description="A body paragraph formatted to look like a heading.",
    tags={"structure", "headings"},
)
def manual_heading_formatting(ctx: LintContext) -> Iterator[Issue]:
    """Flag body paragraphs bolded or enlarged to stand in for a heading.

    These are invisible to the navigation pane, absent from every generated
    table of contents, and unreachable by an accessibility tool — the
    document *looks* structured and is not.

    Deliberately conservative, because the shape (a short bold line) also
    describes plenty of legitimate content. Every signal must hold: the
    paragraph is unstyled body text, has text, is short, does not end like
    a sentence, is not a list item, is not in a table cell — table cells
    are full of short bold labels — and is either wholly bold or set larger
    than the document's own body size. That last comparison is what makes
    the rule a document-relative judgement rather than an opinion: "larger
    than this document's body text", not "larger than 11pt".

    **Unstyled** is doing real work there. A paragraph carrying any
    non-default style has been styled deliberately, whatever it looks like;
    it is not a body paragraph dressed up as a heading. The template's
    ``Caption`` style is bold, so without this the rule reports every
    caption in the document.

    Report-only: applying a heading style needs a *level*, and the fact
    that a line is bold says nothing about how deep it sits.
    """
    body_size = _dominant_body_size(ctx)
    default_style = _default_paragraph_style(ctx)

    for resolved in ctx.paragraphs:
        if resolved.in_table or resolved.formatting.num_id:
            continue
        style_id = resolved.formatting.style_id
        if style_id is not None and style_id != default_style:
            continue

        text = resolved.text.strip()
        if not text or len(text) > _HEADING_LIKE_MAX_CHARS or text[-1] in ".?!:;,":
            continue

        runs = [r for r in resolved.runs if r.run.text.strip()]
        if not runs:
            continue

        # Read the size off the runs, not the paragraph: a paragraph resolve
        # reports its mark's formatting, and a line enlarged by hand almost
        # always carries the size on the runs.
        sizes = [r.formatting.font_size for r in runs if r.formatting.font_size is not None]
        smallest = min(sizes) if len(sizes) == len(runs) else None

        if all(r.formatting.bold for r in runs):
            signal = "bold"
        elif body_size is not None and smallest is not None and smallest > body_size:
            signal = f"{smallest}pt against the document's {body_size}pt body text"
        else:
            continue

        yield Issue(
            message=(
                f"Paragraph is formatted like a heading ({signal}) but carries no "
                f"heading style, so it is invisible to the outline and any TOC."
            ),
            location=Location(
                paragraph_index=resolved.index,
                style_id=resolved.formatting.style_id,
                excerpt=ctx.excerpt(resolved.index),
            ),
            observed=signal,
            expected="a heading style",
        )


def _default_paragraph_style(ctx: LintContext) -> str | None:
    """The id of the document's default paragraph style, usually ``Normal``.

    Read from the document rather than assumed, since a template is free to
    name its default anything.
    """
    for info in list_styles(ctx.doc, style_type="paragraph"):
        if info.is_default:
            return info.style_id
    return None


def _dominant_body_size(ctx: LintContext) -> float | None:
    """The most common resolved run size outside headings.

    The document's own answer to "how big is body text", so the comparison
    holds for a 10pt house style as readily as a 12pt one. Counted per run
    rather than per paragraph, since that is where a size usually sits.
    """
    counts = Counter(
        run.formatting.font_size
        for resolved in ctx.paragraphs
        if _heading_level(resolved) is None
        for run in resolved.runs
        if run.run.text.strip() and run.formatting.font_size is not None
    )
    if not counts:
        return None
    return counts.most_common(1)[0][0]
