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
from docx_plus.lint.models import Issue, Location
from docx_plus.lint.registry import rule

if TYPE_CHECKING:
    from collections.abc import Iterator

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

_DOUBLE_SPACE = re.compile(r"\S {2,}\S")
_SPACE_BEFORE_PUNCT = re.compile(r"\s+([,.;:!?])")


def _is_verbatim(resolved: ResolvedParagraph) -> bool:
    """True for styles where runs of whitespace are content, not sloppiness."""
    return (resolved.formatting.style_id or "") in _VERBATIM_STYLES


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
    """Flag runs of 2+ spaces between words."""
    for resolved in ctx.paragraphs:
        if _is_verbatim(resolved):
            continue
        if _DOUBLE_SPACE.search(resolved.text):
            yield Issue(
                message="Two or more consecutive spaces between words.",
                location=Location(
                    paragraph_index=resolved.index,
                    excerpt=ctx.excerpt(resolved.index),
                ),
                observed="multiple spaces",
                expected="one space",
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
            yield Issue(
                message="Paragraph ends in trailing whitespace.",
                location=Location(
                    paragraph_index=resolved.index,
                    excerpt=ctx.excerpt(resolved.index),
                ),
                observed=f"ends with {text[len(text.rstrip()) :]!r}",
                expected="no trailing whitespace",
            )


@rule(
    id="space-before-punctuation",
    kind="consistency",
    severity="info",
    description="Whitespace before a comma, period, semicolon, colon, or other closing mark.",
    tags={"typography"},
)
def space_before_punctuation(ctx: LintContext) -> Iterator[Issue]:
    """Flag a space sitting before closing punctuation."""
    for resolved in ctx.paragraphs:
        if _is_verbatim(resolved):
            continue
        match = _SPACE_BEFORE_PUNCT.search(resolved.text)
        if match is not None:
            yield Issue(
                message=f"Whitespace before {match.group(1)!r}.",
                location=Location(
                    paragraph_index=resolved.index,
                    excerpt=ctx.excerpt(resolved.index),
                ),
                observed=match.group(0).replace("\t", "\\t"),
                expected=match.group(1),
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

    The eventual fix deletes content, so it carries ``adds_content``.
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
            yield Issue(
                message=f"{run_length} consecutive empty paragraphs used as spacing.",
                location=Location(paragraph_index=run_start),
                observed=f"{run_length} empty paragraphs",
                expected="paragraph spacing (w:spacing)",
                adds_content=True,
            )
        run_start = None
        run_length = 0
