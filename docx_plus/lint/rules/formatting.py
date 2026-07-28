"""Direct-formatting consistency rules — the linter's centre of gravity.

Professional documents are style-driven, and the defects people actually
spend time undoing are almost always *direct overrides that drifted from
the style*. So most of "make this consistent" reduces to one question:
**does this run carry a direct property that the cascade would have given
it anyway, or that fights what the style says?**

This is where the OOXML approach beats a COM one. Word's object model
exposes an effective value and the applied paragraph style — a two-layer
compare that cannot say *which* layer set a property, and offers no
per-property reset. Here, ``FormattingSource`` names the exact layer that
won, and resolving the same target with ``stop_below`` says what would
have surfaced without it. Between them, a rule can tell "redundant, delete
it" from "this genuinely overrides the style" from "a character style did
this, leave it alone".
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx_plus.lint.models import Issue, Location
from docx_plus.lint.registry import rule

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator

    from docx_plus.lint.models import LintContext
    from docx_plus.styles import ResolvedFormatting


# The run properties worth comparing. Deliberately not every field on
# ResolvedFormatting: paragraph-level properties do not belong to a run,
# and `style_id` / `style_name` are identity, not formatting.
_RUN_PROPERTIES = (
    "font_name",
    "font_size",
    "bold",
    "italic",
    "underline",
    "strike",
    "double_strike",
    "color_rgb",
    "highlight",
    "caps",
    "small_caps",
    "vert_align",
)

# Boolean toggles, where an unset property and an explicit "off" render
# identically. Everything else distinguishes None (inherit / renderer
# default) from a concrete value, so None must not be normalised there.
_TOGGLE_PROPERTIES = frozenset({"bold", "italic", "strike", "double_strike", "caps", "small_caps"})

# The paragraph properties `style-drift` compares. Numbering is excluded:
# `num_id` has its own rule, and a paragraph's indent legitimately comes
# from the numbering level rather than its style.
_PARAGRAPH_PROPERTIES = (
    "alignment",
    "indent_left",
    "indent_right",
    "indent_first_line",
    "spacing_before",
    "spacing_after",
    "line_spacing",
    "line_spacing_rule",
    "keep_with_next",
    "keep_lines",
    "page_break_before",
)

_LABELS = {
    "font_name": "font",
    "font_size": "size",
    "color_rgb": "colour",
    "double_strike": "double strikethrough",
    "small_caps": "small caps",
    "vert_align": "vertical alignment",
    "indent_left": "left indent",
    "indent_right": "right indent",
    "indent_first_line": "first-line indent",
    "spacing_before": "space before",
    "spacing_after": "space after",
    "line_spacing": "line spacing",
    "line_spacing_rule": "line spacing rule",
    "keep_with_next": "keep with next",
    "keep_lines": "keep lines together",
    "page_break_before": "page break before",
}


def _label(prop: str) -> str:
    return _LABELS.get(prop, prop.replace("_", " "))


def _same_rendering(prop: str, direct: object, inherited: object) -> bool:
    """Whether a direct value renders identically to the inherited one.

    For a toggle, ``None`` and ``False`` are the same picture — an unset
    ``w:b`` and an explicit ``<w:b w:val="0"/>`` both draw upright text — so
    writing "not bold" over nothing is redundant even though the resolved
    values differ. That case is common enough to matter: it is what
    select-all-then-clear-formatting leaves behind.
    """
    if prop in _TOGGLE_PROPERTIES:
        return bool(direct) == bool(inherited)
    return direct == inherited


def _set_directly(
    formatting: ResolvedFormatting,
    properties: Iterable[str],
    layer: str,
) -> list[str]:
    """The properties whose winning value came from ``layer``."""
    provenance = formatting.provenance or {}
    return [
        prop
        for prop in properties
        if (source := provenance.get(prop)) is not None and source.layer == layer
    ]


@rule(
    id="redundant-direct-formatting",
    kind="consistency",
    severity="info",
    description="A run sets a property directly to the value it would inherit anyway.",
    tags={"formatting", "styles"},
)
def redundant_direct_formatting(ctx: LintContext) -> Iterator[Issue]:
    """Flag direct run formatting that changes nothing.

    The classic select-all-and-set-the-font mess: every run carries an
    explicit ``w:sz`` and ``w:rFonts`` identical to what the style already
    said. It renders correctly, so nobody notices — until the style is
    changed and nothing moves, because the direct formatting overrides it.

    Detection compares each run's resolved value against its ``baseline`` —
    the same run resolved with ``stop_below="directRun"``, which is exactly
    "what this run would render as if its own ``<w:rPr>`` were deleted".
    That baseline still includes the run's character style, so a property
    a ``w:rStyle`` supplies is correctly *not* called redundant, and a run
    carrying one is checked like any other.
    """
    for resolved in ctx.paragraphs:
        for resolved_run in resolved.runs:
            baseline = resolved_run.baseline
            if baseline is None:
                continue

            redundant = [
                prop
                for prop in _set_directly(resolved_run.formatting, _RUN_PROPERTIES, "directRun")
                if _same_rendering(
                    prop,
                    getattr(resolved_run.formatting, prop),
                    getattr(baseline, prop),
                )
            ]
            if not redundant:
                continue

            names = ", ".join(_label(p) for p in redundant)
            yield Issue(
                message=(
                    f"Run sets {names} directly to the value it already inherits; "
                    f"the direct formatting has no effect but overrides the style."
                ),
                location=Location(
                    paragraph_index=resolved.index,
                    run_index=resolved_run.index,
                    style_id=resolved.formatting.style_id,
                    excerpt=ctx.excerpt(resolved.index),
                ),
                observed=", ".join(
                    f"{_label(p)}={getattr(resolved_run.formatting, p)!r}" for p in redundant
                ),
                expected="inherited from the style",
            )


@rule(
    id="style-drift",
    kind="consistency",
    severity="warning",
    description="A paragraph's direct formatting overrides what its style says.",
    tags={"formatting", "styles"},
)
def style_drift(ctx: LintContext) -> Iterator[Issue]:
    """Flag paragraph-level direct formatting that fights the applied style.

    The counterpart to ``redundant-direct-formatting``: that rule catches a
    direct value the cascade would have supplied anyway, this one catches a
    direct value that genuinely differs. Both are the same comparison
    against the same baseline, split because the two say opposite things
    about what to do — a redundant property can simply be deleted, while a
    drifted one is a decision: was the nudge deliberate, or should the
    style change?

    Scoped to **paragraph-level** properties. Run-level drift is nearly
    always deliberate — a bolded term, an emphasised phrase — whereas an
    individually nudged indent or space-after is the classic accumulated
    mess: the document looks consistent until the style is edited and forty
    paragraphs refuse to move.

    This is the sibling COM linter's central rule, and the place the OOXML
    approach is strictly better. Word's object model can compare a
    paragraph against its style but cannot say *which* layer produced the
    effective value, so a value arriving from the numbering level or a
    table style reads as drift. Resolving the cascade ourselves, the
    baseline is the real one.
    """
    for resolved in ctx.paragraphs:
        baseline = resolved.baseline
        if baseline is None:
            continue

        drifted = [
            prop
            for prop in _set_directly(resolved.formatting, _PARAGRAPH_PROPERTIES, "directParagraph")
            if not _same_rendering(
                prop,
                getattr(resolved.formatting, prop),
                getattr(baseline, prop),
            )
        ]
        if not drifted:
            continue

        style = resolved.formatting.style_name or resolved.formatting.style_id or "the style"
        names = ", ".join(_label(p) for p in drifted)
        yield Issue(
            message=f"Paragraph overrides {names} directly, deviating from {style}.",
            location=Location(
                paragraph_index=resolved.index,
                style_id=resolved.formatting.style_id,
                excerpt=ctx.excerpt(resolved.index),
            ),
            observed=", ".join(f"{_label(p)}={getattr(resolved.formatting, p)!r}" for p in drifted),
            expected=", ".join(f"{_label(p)}={getattr(baseline, p)!r}" for p in drifted),
        )


@rule(
    id="mixed-run-formatting",
    kind="consistency",
    severity="info",
    description="Runs within one paragraph disagree on font or size.",
    tags={"formatting"},
    default_on=False,
)
def mixed_run_formatting(ctx: LintContext) -> Iterator[Issue]:
    """Flag a paragraph whose runs resolve to different fonts or sizes.

    Off by default: mixed formatting within a paragraph is frequently
    deliberate (an inline code span, an emphasised term). It earns its
    place as an opt-in because the *accidental* version — a sentence
    pasted in at 11.5pt among 11pt text — is invisible until printed.

    Worth noting this rule is report-only in the sibling COM linter,
    because Word returns an "undefined" sentinel for a paragraph whose runs
    disagree and cannot say which run is the outlier. Sweeping the runs
    ourselves means we can name them.
    """
    for resolved in ctx.paragraphs:
        if len(resolved.runs) < 2:
            continue
        for prop in ("font_name", "font_size"):
            values = {
                getattr(r.formatting, prop)
                for r in resolved.runs
                if r.run.text.strip() and getattr(r.formatting, prop) is not None
            }
            if len(values) > 1:
                rendered = ", ".join(sorted(repr(v) for v in values))
                yield Issue(
                    message=f"Runs in this paragraph use different {_label(prop)} values.",
                    location=Location(
                        paragraph_index=resolved.index,
                        excerpt=ctx.excerpt(resolved.index),
                    ),
                    observed=rendered,
                )
