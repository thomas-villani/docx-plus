"""Rules about the style *definitions* rather than the content.

The first rules here whose subject is not a paragraph. Findings carry a
``style_id`` and no position, which is why every field on
:class:`~docx_plus.lint.models.Location` is optional.

This is also where the composing-layer design earns itself: the OOXML
knowledge — what counts as a reference to a style, how to walk the
``basedOn`` chain — lives in ``styles/``, and the rules here only decide
what is worth reporting.
"""

from __future__ import annotations

from collections import defaultdict
from typing import TYPE_CHECKING

from docx_plus.lint.models import Fix, FixOperation, Issue, Location
from docx_plus.lint.registry import rule
from docx_plus.styles import find_unused_styles

if TYPE_CHECKING:
    from collections.abc import Iterator

    from docx_plus.lint.models import LintContext
    from docx_plus.styles import ResolvedFormatting


# Identity, not formatting — two styles differing only in these fields are
# exactly the duplicates this rule is looking for.
_IDENTITY_FIELDS = ("style_id", "style_name", "provenance", "partial")


def _formatting_key(formatting: ResolvedFormatting) -> tuple[tuple[str, object], ...]:
    """A comparable snapshot of what a resolve *renders as*."""
    return tuple(
        (name, value)
        for name, value in sorted(vars(formatting).items())
        if name not in _IDENTITY_FIELDS
    )


@rule(
    id="duplicate-styles",
    kind="consistency",
    severity="info",
    description="Two or more styles that resolve to identical formatting.",
    tags={"styles"},
)
def duplicate_styles(ctx: LintContext) -> Iterator[Issue]:
    """Flag style ids that render identically.

    What "Keep Source Formatting" leaves behind: ``Body Text``,
    ``BodyText1``, and ``Default Paragraph`` all resolving to the same
    thing, so editing the one you meant moves a third of the document.

    Compares **resolved** formatting, not the style elements, so two styles
    reaching the same place by different ``basedOn`` routes still match —
    and it compares only styles the document actually *uses*, since a
    duplicate nobody applied is ``unused-styles``' finding instead.

    The comparison uses each paragraph's ``baseline``, which is the resolve
    with the paragraph's own direct formatting excluded: docDefaults plus
    the style chain, and nothing the author typed on top. Numbered and
    table paragraphs are skipped, since their baselines also carry the
    numbering level and the table style — neither of which belongs to the
    paragraph style being compared.

    Report-only. Merging duplicates means remapping every paragraph onto
    one of them and deleting the rest, and *which* one survives is not
    something the document answers: the three ids render alike, so nothing
    in the formatting favours any of them. Picking by name or by first use
    would be the plan making a decision it cannot justify.
    """
    by_formatting: defaultdict[tuple[tuple[str, object], ...], dict[str, int]] = defaultdict(dict)

    for resolved in ctx.paragraphs:
        style_id = resolved.formatting.style_id
        baseline = resolved.baseline
        if style_id is None or baseline is None:
            continue
        if resolved.in_table or resolved.formatting.num_id:
            continue
        # First paragraph wins as the exemplar, so the report points at the
        # earliest occurrence of each style.
        by_formatting[_formatting_key(baseline)].setdefault(style_id, resolved.index)

    for styles in by_formatting.values():
        if len(styles) < 2:
            continue
        names = ", ".join(sorted(styles))
        for style_id, paragraph_index in sorted(styles.items()):
            others = sorted(other for other in styles if other != style_id)
            yield Issue(
                message=(
                    f"Style {style_id} resolves identically to "
                    f"{', '.join(others)}; editing one will not move the rest."
                ),
                location=Location(
                    paragraph_index=paragraph_index,
                    style_id=style_id,
                    excerpt=ctx.excerpt(paragraph_index),
                ),
                observed=f"{len(styles)} styles resolving alike: {names}",
            )


@rule(
    id="unused-styles",
    kind="structural",
    severity="info",
    description="An author-created style is defined but referenced nowhere.",
    tags={"styles"},
    default_on=False,
)
def unused_styles(ctx: LintContext) -> Iterator[Issue]:
    """Flag author-created style definitions nothing refers to.

    **Word's built-ins are excluded, and that is the whole rule.** A stock
    template materialises over 160 style definitions and a typical document
    uses a handful, so reporting every unreferenced one buries a real
    finding under the entire style gallery — measured at 165 findings on an
    empty two-paragraph document. Those definitions are what the template
    is *for*; they are not a defect. What is worth reporting is the style
    someone made for one heading and abandoned.

    Usage is a **closure**, not a single pass — see
    :func:`~docx_plus.styles.find_unused_styles`. A style referenced only
    by another unused style is itself unused, and a style reached through a
    used style's ``basedOn`` is not.

    Off by default, and ``info`` when on: an unused style is untidiness
    rather than a defect, and a document deliberately carrying styles for
    later use is exactly this shape.
    """
    for info in find_unused_styles(ctx.doc):
        if info.is_builtin:
            continue
        yield Issue(
            message=(
                f"Style {info.style_id} ({info.name}) is defined but nothing in "
                f"the document refers to it."
            ),
            location=Location(style_id=info.style_id),
            observed=f"{info.style_type} style, unreferenced",
            # Removing a style definition is a content change, not a
            # formatting one — nothing renders differently, but the
            # document loses something it had.
            adds_content=True,
            fix=Fix(
                summary=f"Delete the unused style {info.style_id}.",
                safety="destructive",
                operations=(FixOperation(op="delete-style", args={"style_id": info.style_id}),),
            ),
        )
