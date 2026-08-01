# Styles: remapping and built-ins

Two things sit between "I want a Heading 1" and a document that actually
has one: the style id may not be the one you expect, and the style may not
be defined at all. This page covers both. The [cascade
resolver](cascade.md) covers the read side; the [styles
guide](../guides/styles.md) covers the calls.

## Style remapping

Real-world documents have a long-running mismatch between style IDs (the
`w:styleId` attribute, what code references) and style names (the
`w:name` attribute, what Word shows in its UI). The same logical style
might appear as `Heading1` in one doc, `Heading 1` (with space) in
another, `heading1` in a third, and `HeadingOne` in a fourth. Code that
calls `apply_style(p, "Heading1")` against the second doc fails — not
because the style is missing, but because the ID doesn't match.

`styles/modify.py:find_matching_style` (line 550) does case- and
space-insensitive lookup against both `w:styleId` and `w:name` of every
defined style. It returns the trivial match when the exact ID is
defined, so it is safe to call unconditionally.

`styles/modify.py:remap_styles` (line 585) is the bulk reconciliation.
For each target ID it walks four steps:

1. **Exact match** — if `target_id` is already defined as a `w:styleId`,
   record the trivial mapping and continue
2. **Explicit mapping** — if `mapping[target_id]` is in the caller's
   dict and points at an existing style, use it
3. **Matcher** — call `find_matching_style(doc, target_id)`. If a
   case/space-insensitive match exists, use the existing definition
4. **Create from built-ins** — only if `create_missing=True` and the
   target is in `_BUILTIN_STYLES`, materialise it via
   `_materialise_builtin`

After resolution, body references (`w:pStyle`, `w:rStyle`, `w:tblStyle`)
are rewritten in-place so subsequent `apply_style` calls work without
translation.

Style-to-style references inside `styles.xml` (`w:basedOn`, `w:next`,
`w:link`, `w:numStyleLink`, `w:styleLink`) are **intentionally not
rewritten**. The remap is a non-destructive body-only rewrite — if the
authoring tool chained `MyHeading` as `basedOn="HeadingOne"`, the chain
is preserved. The cascade resolver will follow it correctly because the
matcher feeds the `apply_style` path, not the inheritance walker.

`ensure_style` accepts a `match_existing=False` flag (added in Phase
3.5). With `True`, it consults `find_matching_style` before falling
back to the built-ins / custom-create path. The returned proxy's
`style_id` may differ from the requested one — callers using
`apply_style` should pass `proxy.style_id` or use `remap_styles` for
document-wide normalisation.

## The built-in styles table

`_BUILTIN_STYLES` in `styles/modify.py:1154` enumerates **107 of Word's
built-in styles** — well past SPEC §5's "at minimum" set, covering
essentially every style a real Word user reaches for. The entries are
grouped into seven tiers:

| Tier | Count | Coverage |
|---|---:|---|
| Core | 19 | `Normal`, `Heading1`–`Heading9`, `Title`, `Subtitle`, `Quote`, `IntenseQuote`, `ListParagraph`, `Caption`, `DefaultParagraphFont`, `Hyperlink`, `PlaceholderText`, `TableNormal`, `NoList` |
| A — structural essentials | 6 | `NoSpacing`, `Header`/`HeaderChar`, `Footer`/`FooterChar`, `TableGrid` |
| B — character emphasis | 7 | `Strong`, `Emphasis`, `IntenseEmphasis`, `SubtleEmphasis`, `IntenseReference`, `SubtleReference`, `BookTitle` |
| C — heading linked-Char | 13 | `Heading1Char`–`Heading9Char`, `TitleChar`, `SubtitleChar`, `QuoteChar`, `IntenseQuoteChar` |
| D — lists | 18 | `List`/`List2`/`List3`, `ListBullet`/`2`–`5`, `ListNumber`/`2`–`5`, `ListContinue`/`2`–`5` |
| E — TOC / index / table-of-* | 16 | `TOCHeading`, `TOC1`–`TOC9`, `IndexHeading`, `Index1`, `TableofFigures`, `TableofAuthorities`, `TOAHeading` |
| F — footnotes / endnotes / comments | 12 | `FootnoteText`/`Char`, `FootnoteReference`, `EndnoteText`/`Char`, `EndnoteReference`, `CommentText`/`Char`, `CommentReference`, `CommentSubject`/`Char`, `BalloonText`/`Char` |
| G — body / macro / preformatted | 16 | `BodyText`/`2`/`3` + Char companions, `MacroText`/`Char`, `HTMLPreformatted`/`Char`, `PlainText`/`Char`, `NormalIndent`, `BlockText` |

Defaults come from extracting `styles.xml` from real Word-saved
documents (Word 365, 2026-05-19) — *not* from guessing or copying
Word-2007 numbers. About 65 entries (Core, A, B, the most-common subset
of C–G) are sourced from python-docx's bundled `default.docx`; the
latent remainder (TOC*, footnote/endnote/comment family, Index*, table-
of-*, HTMLPreformatted, PlainText, BodyText, MacroText, BalloonText,
BlockText) were extracted from Word-saved sample docs that materialise
each style after it's applied to a paragraph.

Built-ins materialise *without* `w:customStyle="1"` (they are not
user-defined) and the four `default` entries carry `w:default="1"`.

### Known property-writer limitations

A handful of Word's defaults can't currently be emitted because the
property writer doesn't model them — these are intentionally omitted from
`_BUILTIN_STYLES`:

- **Theme attributes** (`themeColor`, `themeShade`, `asciiTheme`, etc.)
  on `Heading*Char`, `Caption`, `IntenseQuote`, `IndexHeading`,
  `TOAHeading`. Literal RGB/font values are emitted instead — visually
  equivalent for users on Word's default Office theme.
- **`semiHidden` / `unhideWhenUsed`** presence-only flags on latent
  styles. Not a property kind we expose; styles still work, they just
  always show in Word's style gallery.
- **Tab stops** on `Header`, `Footer`, `MacroText`.
- **Paragraph borders** (`pBdr`) on `IntenseQuote`, `BlockText`.
- **`numPr` placeholder** on `ListBullet`/`ListNumber` — these styles
  in Word's default ship with an empty `numPr` child (a hint, no real
  numbering link). Skipped; callers attach numbering separately.

### `ensure_style` never overwrites

`ensure_style` is idempotent and aware that **python-docx already ships
a `styles.xml` with many of these latent built-ins materialised** at
Word-2007 defaults (e.g. Heading1 = 14pt #365F91), not Word-2013/365.
This is deliberate: `ensure_style` consults the built-ins table **only**
when the ID is genuinely missing from `styles.xml`. If python-docx
already shipped it, the existing definition is returned unchanged. The
table is a "the style is absent, here is what Word would have written"
fallback, not a "force my preferred defaults" mechanism — for that,
use `modify_style` or `remap_styles`.
