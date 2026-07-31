"""Text-hygiene rules — the cheap, high-frequency cleanup pass.

Pure text scans over the swept paragraphs. These need no cascade
information at all, which makes them the natural place to prove the rule
shape end to end.

Whitespace-bearing styles are skipped throughout: a run of spaces inside a
code or preformatted paragraph is content, not a defect.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from docx_plus.core.ns import qn
from docx_plus.lint.models import Fix, FixOperation, Issue, Location
from docx_plus.lint.registry import rule

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator

    from docx_plus.lint.models import LintContext
    from docx_plus.styles import ResolvedParagraph


_VERBATIM_STYLES = frozenset(
    {
        "HTMLPreformatted",
        "PlainText",
        "MacroText",
        "Code",
        "SourceCode",
    }
)

# Lookarounds rather than literal ``\S {2,}\S``: the flanking characters
# must not be consumed, or the word between two double spaces is eaten by
# the first match and the second goes unreported. Matching only the spaces
# also makes the match span exactly the span the fix replaces.
_DOUBLE_SPACE = re.compile(r"(?<=\S) {2,}(?=\S)")
_SPACE_BEFORE_PUNCT = re.compile(r"\s+([,.;:!?])")


def _is_verbatim(resolved: ResolvedParagraph) -> bool:
    """True for styles where runs of whitespace are content, not sloppiness."""
    return (resolved.formatting.style_id or "") in _VERBATIM_STYLES


def _rewrite(
    paragraph_index: int,
    spans: Iterable[tuple[int, int, str]],
    *,
    summary: str,
) -> Fix:
    """A text fix, as half-open character spans into the paragraph's text.

    Every span is measured against the **original** text rather than
    against the text as the previous span left it, so the spans are
    independent: applying them in any order gives the same result, and two
    rules' spans can be tested for overlap without replaying either.

    That independence is what lets a paragraph carrying a double space
    *and* a space before a comma be fixed by both rules in one pass — the
    common case, and one a "find this, replace it with that" fix model
    cannot express without the two edits fighting.
    """
    return Fix(
        summary=summary,
        safety="review",
        operations=(
            FixOperation(
                op="replace-paragraph-text",
                args={
                    "paragraph_index": paragraph_index,
                    "spans": [
                        {"start": start, "end": end, "replacement": replacement}
                        for start, end, replacement in spans
                    ],
                },
            ),
        ),
    )


def _plural(count: int, noun: str) -> str:
    """``"1 run"`` / ``"3 runs"`` — ASCII, for a cp1252 console."""
    return f"{count} {noun}" if count == 1 else f"{count} {noun}s"


def _has_field(resolved: ResolvedParagraph) -> bool:
    """True if the paragraph contains a complex field.

    Its rendered text is then not its stored text: the field contributes
    its *cached result*, which is whatever Word last displayed.
    """
    return resolved.paragraph._p.find(f".//{qn('w:fldChar')}") is not None


@rule(
    id="double-space",
    kind="consistency",
    severity="info",
    description="Two or more spaces between words in body text.",
    tags={"typography", "whitespace"},
)
def double_space(ctx: LintContext) -> Iterator[Issue]:
    """Flag runs of 2+ spaces between words.

    One finding per paragraph however many times it occurs — a paragraph
    with four double spaces has one problem, not four — but the fix carries
    every occurrence, so repairing the finding repairs the paragraph.
    """
    for resolved in ctx.paragraphs:
        if _is_verbatim(resolved):
            continue
        spans = [
            (match.start(), match.end(), " ") for match in _DOUBLE_SPACE.finditer(resolved.text)
        ]
        if spans:
            yield Issue(
                message="Two or more consecutive spaces between words.",
                location=Location(
                    paragraph_index=resolved.index,
                    excerpt=ctx.excerpt(resolved.index),
                ),
                observed="multiple spaces",
                expected="one space",
                fix=_rewrite(
                    resolved.index,
                    spans,
                    summary=f"Collapse {_plural(len(spans), 'run')} of spaces to a single space.",
                ),
            )


@rule(
    id="trailing-whitespace",
    kind="structural",
    severity="info",
    description="A paragraph ending in a space or tab.",
    tags={"typography", "whitespace"},
)
def trailing_whitespace(ctx: LintContext) -> Iterator[Issue]:
    """Flag paragraphs whose text ends in whitespace.

    Paragraphs containing a field are skipped. A field's text is its
    *cached result* — whatever Word last rendered, often empty in a
    freshly-written document — so ``"See "`` followed by an unrendered
    ``REF`` reads as trailing whitespace when the space is doing exactly
    its job. Every cross-reference and page number would otherwise report.
    """
    for resolved in ctx.paragraphs:
        text = resolved.text
        if _is_verbatim(resolved) or not text.strip() or _has_field(resolved):
            continue
        if text != text.rstrip():
            stripped = len(text.rstrip())
            yield Issue(
                message="Paragraph ends in trailing whitespace.",
                location=Location(
                    paragraph_index=resolved.index,
                    excerpt=ctx.excerpt(resolved.index),
                ),
                observed=f"ends with {text[stripped:]!r}",
                expected="no trailing whitespace",
                fix=_rewrite(
                    resolved.index,
                    [(stripped, len(text), "")],
                    summary=f"Delete {_plural(len(text) - stripped, 'trailing character')}.",
                ),
            )


@rule(
    id="space-before-punctuation",
    kind="consistency",
    severity="info",
    description="Whitespace before a comma, period, semicolon, colon, or other closing mark.",
    tags={"typography"},
)
def space_before_punctuation(ctx: LintContext) -> Iterator[Issue]:
    """Flag a space sitting before closing punctuation.

    Reported once per paragraph, on the first occurrence, but the fix
    carries every one — same split as ``double-space``.
    """
    for resolved in ctx.paragraphs:
        if _is_verbatim(resolved):
            continue
        matches = list(_SPACE_BEFORE_PUNCT.finditer(resolved.text))
        if matches:
            first = matches[0]
            yield Issue(
                message=f"Whitespace before {first.group(1)!r}.",
                location=Location(
                    paragraph_index=resolved.index,
                    excerpt=ctx.excerpt(resolved.index),
                ),
                observed=first.group(0).replace("\t", "\\t"),
                expected=first.group(1),
                fix=_rewrite(
                    resolved.index,
                    [(match.start(), match.start(1), "") for match in matches],
                    summary=(
                        f"Remove the whitespace before {_plural(len(matches), 'punctuation mark')}."
                    ),
                ),
            )


@rule(
    id="indent-by-whitespace",
    kind="structural",
    severity="warning",
    description="Leading spaces or tabs standing in for a real indent.",
    tags={"typography", "whitespace"},
)
def indent_by_whitespace(ctx: LintContext) -> Iterator[Issue]:
    """Flag paragraphs indented with typed whitespace rather than ``w:ind``.

    Typed indentation does not survive a style change, a column change, or
    translation into another layout, which is the whole reason indents are
    a paragraph property.

    Report-only, and the reason is worth stating because it looks fixable.
    Deleting the whitespace is half a repair: the paragraph then has no
    indent at all, which is not what the author wanted either. The other
    half needs a number, and how far four typed spaces were meant to indent
    is not recoverable from the document — a plan supplying one would be
    the library inventing a house style, which is the thing the rule kinds
    exist to prevent.
    """
    for resolved in ctx.paragraphs:
        text = resolved.text
        if _is_verbatim(resolved) or not text.strip():
            continue
        leading = text[: len(text) - len(text.lstrip(" \t"))]
        if "\t" in leading or len(leading) >= 2:
            yield Issue(
                message="Paragraph is indented with typed whitespace, not an indent property.",
                location=Location(
                    paragraph_index=resolved.index,
                    excerpt=ctx.excerpt(resolved.index),
                ),
                observed=f"{len(leading)} leading whitespace characters",
                expected="a w:ind indent",
            )


@rule(
    id="stray-empty-paragraph",
    kind="structural",
    severity="info",
    description="Consecutive empty paragraphs used as vertical spacing.",
    tags={"typography", "whitespace"},
    default_on=False,
)
def stray_empty_paragraph(ctx: LintContext) -> Iterator[Issue]:
    """Flag runs of 2+ consecutive empty paragraphs.

    Off by default, and deliberately only firing on a *run* of them: a
    single empty paragraph is common and often intentional, while two in a
    row is nearly always spacing that belongs in ``spaceAfter``.

    The fix deletes content, so it carries ``adds_content`` and a plan
    withholds it unless the caller asks for it.
    """
    run_start: int | None = None
    run_length = 0

    for resolved in [*ctx.paragraphs, None]:
        empty = resolved is not None and not resolved.text.strip()
        if empty and resolved is not None:
            if run_start is None:
                run_start = resolved.index
            run_length += 1
            continue
        if run_start is not None and run_length >= 2:
            # One is kept. The finding is "these are doing spacing's job",
            # not "this gap should not exist", and collapsing the run to a
            # single blank line is the smallest change that says so.
            #
            # Emitted back to front, because each deletion shifts every
            # index after it. The plan orders *fixes* that way too, but a
            # fix's own operations are applied in the order it gives them,
            # so a fix that deletes more than one thing has to sequence
            # itself.
            doomed = range(run_start + run_length - 1, run_start, -1)
            yield Issue(
                message=f"{run_length} consecutive empty paragraphs used as spacing.",
                location=Location(paragraph_index=run_start),
                observed=f"{run_length} empty paragraphs",
                expected="paragraph spacing (w:spacing)",
                adds_content=True,
                fix=Fix(
                    summary=(f"Delete {_plural(len(doomed), 'empty paragraph')}, keeping one."),
                    safety="destructive",
                    operations=tuple(
                        FixOperation(op="delete-paragraph", args={"paragraph_index": index})
                        for index in doomed
                    ),
                ),
            )
        run_start = None
        run_length = 0
