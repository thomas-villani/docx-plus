"""Rules about fields — cross-references and captions.

The defects here are the ones that survive proofreading, because the
document *looks* right until something moves. A `REF` to a bookmark that no
longer exists renders as ``Error! Reference source not found.`` only after
Word recalculates; a caption numbered by hand looks perfect until a figure
is inserted above it.

Both read the field instruction rather than the cached result, since the
result is whatever Word last rendered and can be arbitrarily stale — which
is precisely how a broken reference stays invisible.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from docx_plus.bookmarks import read_bookmarks
from docx_plus.core.oxml import xpath
from docx_plus.fields import read_fields
from docx_plus.lint.models import Issue, Location, render_for_report
from docx_plus.lint.registry import rule

if TYPE_CHECKING:
    from collections.abc import Iterator

    from docx_plus.lint.models import LintContext


# Word's own caption labels, plus the ones people type. The number may be
# chapter-qualified ("Figure 2-1", "Table 3.4").
_TYPED_CAPTION = re.compile(
    r"^\s*(Figure|Fig\.?|Table|Equation|Chart|Exhibit|Appendix|Listing|Scheme|Plate)"
    r"\s+(\d+([-.]\d+)*)",
    re.IGNORECASE,
)

_CAPTION_STYLES = frozenset({"Caption", "TableCaption", "FigureCaption"})

_TARGET_LIMIT = 40
"""How much of a bookmark name to quote inside a message. Shorter than an
excerpt's 60 because the name sits mid-sentence."""


@rule(
    id="broken-cross-reference",
    kind="structural",
    severity="error",
    description="A REF or PAGEREF field names a bookmark that does not exist.",
    tags={"references", "fields"},
)
def broken_cross_reference(ctx: LintContext) -> Iterator[Issue]:
    """Flag cross-references pointing at a bookmark nothing defines.

    The only ``error`` in the catalogue, because there is no reading of
    this where the document is fine: Word renders it as ``Error!
    Reference source not found.`` the moment fields recalculate. Until
    then it shows the *cached* result — the text that was correct when the
    bookmark still existed — which is what lets a deleted heading go
    unnoticed through a dozen revisions.

    Reading the instruction rather than that cached result is the whole
    point. Both halves are already covered:
    :func:`~docx_plus.fields.read_fields` gives the reference and
    :func:`~docx_plus.bookmarks.read_bookmarks` gives the targets.

    Report-only, and unusually clearly so: the bookmark this field wanted
    is gone, and nothing left in the document says what it pointed at.
    """
    defined = {bookmark.name for bookmark in read_bookmarks(ctx.doc)}
    sweep_index = _sweep_index_by_body_index(ctx)

    for found in read_fields(ctx.doc):
        if found.keyword not in ("REF", "PAGEREF"):
            continue
        operands = found.operands
        if not operands:
            continue
        target = operands[0]
        if target in defined:
            continue

        yield Issue(
            # Both the target and the instruction are arbitrary document
            # text: a bookmark name can hold anything, and the report goes
            # to a console that may not be UTF-8.
            message=(
                f"{found.keyword} field points at bookmark "
                f'"{render_for_report(target, _TARGET_LIMIT)}", which is '
                f"not defined anywhere in the document."
            ),
            location=Location(
                paragraph_index=sweep_index.get(found.paragraph_index),
                excerpt=render_for_report(found.instruction),
            ),
            observed=target,
            expected="a defined bookmark",
        )


@rule(
    id="caption-manual-numbering",
    kind="structural",
    severity="warning",
    description="A caption is numbered with typed text instead of a SEQ field.",
    tags={"references", "fields", "captions"},
)
def caption_manual_numbering(ctx: LintContext) -> Iterator[Issue]:
    """Flag caption paragraphs whose number was typed rather than generated.

    Typed caption numbers do not renumber when a figure is inserted, and
    nothing can cross-reference them — a ``REF`` needs a bookmark around a
    ``SEQ`` result to have anything to point at. So a document with typed
    captions cannot have working "see Figure 4" references at all, which is
    usually discovered late.

    Scoped to paragraphs already carrying a caption style: a line of body
    text beginning "Table 2" is as likely to be prose about table 2. The
    style resolves through the cascade, so a caption style reached via
    ``basedOn`` still counts.

    Report-only. Swapping typed text for a ``SEQ`` field replaces content
    with a field whose result is whatever Word computes next time, and the
    number it lands on depends on every other caption in the document —
    including the ones this rule is also reporting. That is a whole-document
    renumbering, not a per-paragraph edit.
    """
    sweep_index = _sweep_index_by_body_index(ctx)
    numbered = {
        sweep_index[found.paragraph_index]
        for found in read_fields(ctx.doc, keyword="SEQ")
        if found.paragraph_index in sweep_index
    }

    for resolved in ctx.paragraphs:
        if (resolved.formatting.style_id or "") not in _CAPTION_STYLES:
            continue
        match = _TYPED_CAPTION.match(resolved.text)
        if match is None or resolved.index in numbered:
            continue

        yield Issue(
            message=(
                f"Caption numbers {match.group(1)} {match.group(2)} as literal text; "
                f"it will not renumber, and nothing can cross-reference it."
            ),
            location=Location(
                paragraph_index=resolved.index,
                style_id=resolved.formatting.style_id,
                excerpt=ctx.excerpt(resolved.index),
            ),
            observed=match.group(0).strip(),
            expected=f"a SEQ {match.group(1)} field",
        )


def _sweep_index_by_body_index(ctx: LintContext) -> dict[int, int]:
    """Map ``read_fields``' paragraph numbering onto the sweep's.

    Two numberings meet in this module. The sweep indexes what *it* yields,
    in document order, descending into tables and skipping them entirely
    when asked. ``read_fields`` indexes every ``w:p`` in the body. The two
    diverge at the first table, and diverge differently again under
    ``include_tables=False``.

    Matching by element identity is the only thing that stays correct;
    matching on index would silently mis-attribute a finding to the wrong
    paragraph in any document with a table. Built once per rule run rather
    than per finding.
    """
    sweep_by_element = {resolved.paragraph._p: resolved.index for resolved in ctx.paragraphs}
    return {
        body_index: sweep_by_element[p_element]
        for body_index, p_element in enumerate(xpath(ctx.doc.element.body, ".//w:p"))
        if p_element in sweep_by_element
    }
