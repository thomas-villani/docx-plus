# The cascade resolver

`styles/inspect.py:resolve_effective_formatting` is the algorithmic core of
the library — the thing python-docx most conspicuously lacks. Given a
`Paragraph`, `Run`, or `_Cell`, it walks six layers of OOXML formatting in
precedence order and returns the values that would actually render.

For the task-side view — how to call it, read provenance, and sweep a whole
document — see the [styles guide](../guides/styles.md).

## Six layers, low-to-high precedence

The cascade is walked at `inspect.py:253-317`
(`_apply_paragraph_cascade`):

1. **`docDefaults`** — `w:docDefaults/{w:rPrDefault, w:pPrDefault}` in
   `styles.xml`. Applied by `_apply_doc_defaults` at `inspect.py:337-353`.
2. **Table style** — only if the target is inside a `w:tbl`. For each
   style in the basedOn chain, its base pPr/rPr is applied and then its
   matching **conditional branches** (`w:tblStylePr` for
   firstRow/lastRow/bands/corners), so a child level's base still
   overrides a parent level's conditional. Applied by
   `_apply_table_style_chain`. Which branches match is decided by
   `TableContext` — the cell's position *and* the table's `w:tblLook`
   *and*, for bands, a declared band size. See SPEC "Conditional table
   formatting"; the precedence order deliberately differs from the one
   ECMA-376 17.7.6.5 lists, because Word's does.
3. **Paragraph style chain** — the style named by `w:pStyle` plus every
   `w:basedOn` ancestor. Walked by `_collect_style_chain`, then applied
   root-to-leaf so the most-specific style wins. Cycle detection and depth
   limit (11, per Word) live in that one function. When `w:pStyle` does
   not resolve, `_effective_paragraph_style_id` substitutes the **default
   paragraph style** — see below.
4. **Numbering** — if `w:pPr/w:numPr` is present, the corresponding
   `w:abstractNum/w:lvl` from `numbering.xml` is applied. See
   `_apply_numbering` at `inspect.py:425-466`. If the numbering part is
   missing, `MissingPartError` is **not** raised — the part is treated as
   "not yet materialised" (a common pre-Word state) and skipped silently.
5. **Direct paragraph formatting** — `w:pPr` on the paragraph itself,
   including any `w:rPr` nested under it (paragraph-mark formatting).
6. **Direct run formatting** — `w:rPr` on a target `Run`.

There is deliberately **no linked-character-style layer**. A paragraph
style's `w:link` partner (`Heading1` / `Heading1Char`) is a UI affordance
for applying that style's character half to a selection; Word never
consults it when rendering runs inside the paragraph. `Heading1` carries
its own `<w:b/>`, which is where heading bold comes from.

## The default paragraph style is layer 3, not a second base

Most paragraphs in a real document carry no `w:pStyle` at all, and Word
gives them the default paragraph style — so this is the layer that decides
what an ordinary paragraph looks like. `_ResolverCache.default_paragraph_style_id`
picks it: the **last** `w:type="paragraph"` style whose `w:default` is on,
else the style whose id is `Normal`, else nothing.

It substitutes whenever `w:pStyle` fails to resolve — absent, dangling, or
naming a style of the wrong type — and then behaves as a completely
ordinary paragraph style: it is the reported `style_id`, it supplies
numbering, and it counts as one of the toggle rule's *levels* (a bold
`Normal` and a bold character style cancel).

Its position matters more than its existence. Sitting it under
`docDefaults` would be simpler and would match nearly every measurement,
but it would lose the one case that pins it down: a `Normal` declaring
20pt against a table style declaring 36pt renders at **20pt**. The default
style beats the table style, base and `w:tblStylePr` branches alike, so it
has to be layer 3. `_apply_cell_cascade` therefore ends with it too — a
cell is otherwise just `docDefaults` plus a table style, which is not what
Word reports for an untouched cell.

The other two `w:default="1"` styles do nothing at all:
`DefaultParagraphFont` never reaches a run and `TableNormal` never reaches
a table. Only `w:pStyle` has a fallback.

## Style references are typed

`cache.style(style_id, kind)` and `cache.chain(style_id, kind)` take the
type the *reference* demands, and return nothing when the style is of
another type. This is why `w:rStyle` naming a paragraph style contributes
nothing, and why a `w:basedOn` crossing types ends the chain rather than
extending it. Without it the resolver happily followed all of them.

Both rules were settled against live Word —
`tests/test_default_styles_word_verified.py`.

## Toggle properties

Twelve rPr children are toggles in `_TOGGLE_RPR`: `b`, `i`, `bCs`, `iCs`,
`caps`, `smallCaps`, `strike`, `vanish`, `emboss`, `imprint`, `outline`,
`shadow`. (`dstrike` is intentionally **not** a toggle per ECMA-376.)

Toggle semantics live in `_resolve_toggle`. They are not override, and
they are not a running XOR either — the rule needs one value *per style
level*, which is why `_Accumulator` collects toggles into buckets and
combines them in `freeze()` rather than folding as it walks:

- **`docDefaults` is the base**, not a level.
- **Each style level** — table style, paragraph style, character style —
  is flattened over its own `w:basedOn` chain by **plain override**.
  Inheritance is not a hierarchy boundary.
- The result is the base flipped once per level whose value **differs
  from it**. A level restating the base is inert, so `<w:b w:val="0"/>`
  on a style does nothing when nothing is bold to begin with.
- **Direct formatting is absolute.** `<w:b/>` on a run is bold and
  `<w:b w:val="0"/>` is not, whatever the styles said.

Worked cases:

- Style defines bold, no further override → bold
- Style A bold, B basedOn A bold → **bold** (one level, override)
- Style A bold, B basedOn A `w:b w:val="false"` → not bold (override)
- Paragraph style bold + character style bold → **not bold** (two levels)
- Direct bold on a non-bold style → bold
- Direct bold on a **bold** style → bold (absolute, not a flip)
- Direct `w:b w:val="false"` on a bold style → not bold

This rule was settled by measuring live Word rather than by reading the
spec, whose prose admits several incompatible readings. The measurements
are the parametrised table in `tests/test_cascade_word_verified.py`.

## Paragraph spacing: the one property the cascade cannot finish

Every other field on `ResolvedFormatting` is a pure function of the
cascade. Spacing is not, because `<w:contextualSpacing>` suppresses space
based on the paragraph's **neighbours**. The resolver keeps the two
concerns apart rather than blurring them:

- `spacing_before` / `spacing_after` stay what the cascade declares, and
  `contextual_spacing` carries the flag (a plain override property, not an
  ECMA-376 17.7.3 toggle — `docDefaults` included). Keeping them declared
  is what lets `lint`'s `style-drift` rule compare direct formatting
  against a style, which is a cascade question.
- `resolve_paragraph_spacing(paragraph)` answers the layout question,
  returning a `ParagraphSpacing` whose `space_above` / `space_below` are
  the gaps actually applied. `space_below` of one paragraph equals
  `space_above` of the next by construction.

Both halves of the arithmetic were measured, and both diverged from what
the resolver assumed:

```
gap = after + max(0, before - after)          # i.e. max(after, before)
```

Word lays down the space-after and **tops it up** to the space-before
rather than adding the two. `<w:contextualSpacing>` then removes one of
those two terms — each edge governed only by its own paragraph's flag —
and, crucially, the top-up is still measured from the **declared**
space-after even when that space-after was suppressed. That last detail is
what rules out the obvious "zero the suppressed edge, then take the max".

"Same style" is `styleId` identity: numbering plays no part, and a
`basedOn` child is a different style. Adjacency is sibling adjacency, with
`_adjacent_paragraph` stepping over bookmark and comment markers,
descending into and climbing out of `<w:sdtContent>` (content controls are
transparent), and stopping at a table.

Measured across 111 gaps read out of Word's own layout — COM cannot help
here, since `ParagraphFormat.SpaceBefore` reports the cascade value that
`contextualSpacing` overrides, so the probes were exported to PDF and the
baselines measured. `tests/test_contextual_spacing_word_verified.py` holds
the grid.

## Theme color resolution

Implemented in `styles/theme.py`. `load_theme(doc)` reads
`word/theme/theme1.xml` via the document part's `theme` relationship and
returns a `ThemeColors(scheme=..., fonts=..., mapping=...)`.
`resolve_theme_color(theme, name, *, tint=None, shade=None)` resolves the
`ST_ThemeColor` name to a scheme slot, looks up the base hex, then
applies `themeTint` (toward white) or `themeShade` (toward black).

Name → slot is **per-document**, not a fixed alias table: `settings.xml`
carries a `<w:clrSchemeMapping>` redirecting the semantic names (`text1`,
`background1`, `accent1`, `hyperlink`, …). The direct slot names
(`dark1` / `light1` / `dark2` / `light2`) bypass it. Measured against
Word — a document that remaps `t1` renders `text1` white, and treating
the mapping as fixed resolved it to black.

The transforms use exact rational arithmetic rather than `colorsys`,
because they land on integer boundaries where binary floating point
gives the wrong byte. Verified against Word to within one unit per
channel; `tests/test_theme_word_verified.py` enumerates the residual.

`apply_lum_mod` and `apply_lum_off` implement the DrawingML luminosity
transforms (ECMA-376 17.18.40). They are **not** part of the cascade
path and cannot be: `w:color` has no attribute that carries them, so a
cascade input producing one does not exist. They share this module's
arithmetic but are unverified against Word for the same reason.

Theme failures are **graceful**: if the theme part is missing, malformed,
or names an unknown color, `_resolve_color` at `inspect.py:605-620` sets
`acc.partial = True` and returns the unresolved theme name. The
`ResolvedFormatting.partial` flag tells the caller to expect best-effort
values. SPEC §4 ("Theme resolution edge cases") and
`IMPLEMENTATION.md §5` ("Theme resolution can fail gracefully") both
require this — turning the inspector into something that raises on
diverse real-world inputs would be a usability regression.

## Provenance

When `include_provenance=True`, the resolver populates `ResolvedFormatting.provenance`
with a `FormattingSource` per resolved field. The same walk that produces
values produces provenance (`_Accumulator` carries both, gated on
`want_provenance`); `test_provenance_does_not_change_values` in
`tests/test_cascade_provenance.py` is the regression guard that the
values returned with the flag off are bit-identical to those with it on.

`FormattingSource` records:

- `layer` — which of the eight cascade layers contributed the value
- `style_id` — for `*Style` layers, the lowest style in the basedOn chain
  that actually set the property (not the leaf style, the *resolving*
  style)
- `chain_depth` — how many basedOn hops away from the target
- `is_toggle_resolved` — True when a toggle's value was computed across
  more than one contributing layer rather than stated by one of them

Provenance is the differentiated feature behind the inspector. It is what
makes the [linter](lint.md) possible: every consistency rule is an
assertion about a *layer*, not about a number.
