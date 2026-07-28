"""Direct-formatting consistency rules — the linter's centre of gravity.

Professional documents are style-driven, and the defects people actually
spend time undoing are almost always *direct overrides that drifted from
the style*. So most of "make this consistent" reduces to one question:
**does this run carry a direct property that the cascade would have given
it anyway, or that fights what the style says?**

This is where the OOXML approach beats a COM one. Word's object model
exposes an effective value and the applied paragraph style — a two-layer
compare that cannot say *which* layer set a property, and offers no
per-property reset. Here, ``FormattingSource`` names the exact layer, so a
rule can distinguish "redundant, delete it" from "a character style did
this, leave it alone".
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx_plus.core.ns import qn
from docx_plus.lint.models import Issue, Location
from docx_plus.lint.registry import rule

if TYPE_CHECKING:
    from collections.abc import Iterator

    from docx_plus.lint.models import LintContext
    from docx_plus.styles import ResolvedRun


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

_LABELS = {
    "font_name": "font",
    "font_size": "size",
    "color_rgb": "colour",
    "double_strike": "double strikethrough",
    "small_caps": "small caps",
    "vert_align": "vertical alignment",
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


def _has_run_style(resolved_run: ResolvedRun) -> bool:
    """True if the run carries its own ``w:rStyle`` character style."""
    return resolved_run.run._r.find(f"./{qn('w:rPr')}/{qn('w:rStyle')}") is not None


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

    Detection compares each run's resolved value against the *paragraph's*
    resolved value. The paragraph resolve stops below run-level formatting,
    so it is exactly "what this run would look like without its own
    ``rPr``".

    Runs carrying a ``w:rStyle`` are skipped: for those the paragraph-level
    resolve is not the right baseline, because it also excludes the
    character style, and a property coming from that style is not
    redundant. Catching those needs a resolve beneath only the direct
    layer, which the resolver does not currently offer.
    """
    for resolved in ctx.paragraphs:
        baseline = resolved.formatting
        for resolved_run in resolved.runs:
            provenance = resolved_run.formatting.provenance
            if not provenance or _has_run_style(resolved_run):
                continue

            redundant = [
                prop
                for prop in _RUN_PROPERTIES
                if (source := provenance.get(prop)) is not None
                and source.layer == "directRun"
                and _same_rendering(
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
                    style_id=baseline.style_id,
                    excerpt=ctx.excerpt(resolved.index),
                ),
                observed=", ".join(
                    f"{_label(p)}={getattr(resolved_run.formatting, p)!r}" for p in redundant
                ),
                expected="inherited from the style",
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
