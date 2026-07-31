# `docx_plus.styles.inspect`

The cascade resolver. Walks the OOXML formatting layers and returns a
fully-resolved [`ResolvedFormatting`][docx_plus.styles.inspect.ResolvedFormatting]
plus optional per-field provenance.

See [`ARCHITECTURE.md` §2](../ARCHITECTURE.md#2-the-cascade-resolver) for
the algorithm walkthrough and the toggle semantics.

## Toggle properties do not override

`bold`, `italic` and the ten other ECMA-376 17.7.3 toggles combine by a
rule that catches most people out, so it is worth stating up front:

- A `basedOn` chain is **one level** and flattens by plain override. A
  child style restating its parent's `<w:b/>` stays bold.
- A paragraph style and a character style are **two levels**, and two
  levels both asking for bold cancel. This is the spec's own example.
- `docDefaults` is the **base**, not a level, and a style restating the
  base is inert.
- Direct formatting on a run is **absolute** — it states the value rather
  than flipping it.

The rule was settled by measuring live Word, because the spec prose admits
several incompatible readings and this library previously shipped one of
the wrong ones. `tests/test_cascade_word_verified.py` holds the
measurements.

## A paragraph with no style still has one

Most paragraphs carry no `w:pStyle`, and they are not unstyled — Word gives
them the **default paragraph style**, so `resolve_effective_formatting`
reports it as their `style_id` and applies everything it declares.

- It is chosen as the **last** `w:type="paragraph"` style whose `w:default`
  is on, falling back to the style whose id is `Normal`.
- It substitutes whenever `w:pStyle` fails to resolve — absent, dangling,
  or naming a style of the wrong type. A `w:pStyle` pointing at a style the
  document never defines is reported as the *default* style, not as the
  name it wrote.
- It sits at the paragraph-style layer, so it **beats the table style** —
  a `Normal` declaring 20pt wins over a table style declaring 36pt.
- It is an ordinary toggle *level*: a bold `Normal` plus a bold character
  style cancel.

The default *character* and *table* styles are not the same story — they
never apply to anything. Only `w:pStyle` has a fallback.

A style reference also has to match the style's `w:type`: `w:rStyle` naming
a paragraph style, or `w:basedOn` crossing between the two, contributes
nothing. Measurements: `tests/test_default_styles_word_verified.py`.

## Conditional table formatting is gated, not positional

A table style's `firstRow` / `firstCol` / banding branches do not apply
just because a cell sits in the matching place:

- The table's **`<w:tblLook>`** decides which branches are wanted at all —
  Word's Header Row / First Column / Banded Rows tick-boxes. A cell in row
  0 of a table with `firstRow` cleared takes no `firstRow` formatting, and
  a corner branch needs *both* of its axes enabled. No `<w:tblLook>` at
  all means everything is enabled.
- **Banding needs a declared band size** (`w:tblStyleRowBandSize` or
  `w:tblStyleColBandSize`, on the table or in its style chain). Absent
  means *no banding*, not a band size of one.
- **Precedence is not the order ECMA-376 17.7.6.5 lists.** A vertical band
  beats a horizontal one; a row branch beats a column branch; the corners
  beat everything. `wholeTable` does nothing at all — Word discards the
  branch on load and drops it on save.

Like the toggle rule, this was settled by measuring live Word, because
the spec's prose and Word's behaviour disagree. The measurements live in
`tests/test_tables_word_verified.py`.

## Spacing needs a second call

Spacing is the one property the cascade cannot answer on its own, and the
resolver deliberately does not pretend otherwise.
`ResolvedFormatting.spacing_before` / `spacing_after` are what the cascade
**declares**; `contextual_spacing` carries the resolved
`<w:contextualSpacing>` flag. Whether either value is actually *applied*
depends on the paragraph's neighbours, so
[`resolve_paragraph_spacing`][docx_plus.styles.inspect.resolve_paragraph_spacing]
answers that separately.

Two measured facts it folds in:

- **`<w:contextualSpacing>` drops a paragraph's space on the side where the
  neighbour carries the same `styleId`** — and `styleId` identity is the
  whole test. Numbering plays no part (two `ListParagraph` paragraphs in
  unrelated lists still collapse), and a `basedOn` child counts as a
  different style. Each edge answers only to its own paragraph's flag.
- **Word does not add space-after to the next paragraph's space-before.**
  It tops the first up to the second, so an ordinary pair sits
  `max(after, before)` apart. The top-up is measured from the *declared*
  space-after even when that space-after was suppressed, so a contextual
  paragraph with 20pt after followed by a plain one with 30pt before leaves
  10pt, not 30pt.

A table between two paragraphs stops the suppression; a content control
does not — `<w:sdt>` is transparent. Measurements:
`tests/test_contextual_spacing_word_verified.py`.

## `include_provenance` vs `stop_below`

Two features answer different questions and are easy to confuse:

- **`include_provenance`** names the layer each value *came from*.
- **`stop_below`** re-resolves with a layer excluded, giving the value
  that *would have surfaced without it*.

Provenance alone cannot answer the second: knowing a run's size came from
`directRun` does not say what deleting that `<w:rPr>` would leave. Together
they are what lets a caller tell a direct property that changes nothing
from one that genuinely overrides its style — the basis of every
consistency rule in [`docx_plus.lint`](lint.md).

::: docx_plus.styles.inspect
    options:
      members:
        - resolve_effective_formatting
        - ResolvedFormatting
        - resolve_paragraph_spacing
        - ParagraphSpacing
        - TableContext
        - FormattingSource
        - StyleCascadeError
        - MissingPartError
