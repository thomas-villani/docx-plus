# Styles and the cascade

`docx_plus.styles` has two halves. **Inspection** answers "what does this
actually look like, and why?" — the question python-docx cannot answer,
because most formatting is inherited rather than declared. **Modification**
creates, applies, and changes styles the Word-native way.

## Why is this paragraph 13pt and blue?

```python
from docx import Document
from docx_plus.styles import resolve_effective_formatting

doc = Document("report.docx")
r = resolve_effective_formatting(doc.paragraphs[0])

print(r.style_id, r.style_name)   # "Heading2", "heading 2"
print(r.font_size)                # 13.0   (points)
print(r.bold)                     # True
print(r.color_rgb)                # "2F5496"
```

`None` for a field means "not set at any layer" — it inherits Word's own
default.

Add `include_provenance=True` and each resolved field also tells you which
cascade layer set it:

```python
r = resolve_effective_formatting(doc.paragraphs[0], include_provenance=True)

src = r.provenance["font_size"]
print(src.layer)       # "paragraphStyle"
print(src.style_id)    # "Heading2"  — the style that *resolved* it,
                       #   not necessarily the one applied
print(src.chain_depth) # 0 — how many basedOn hops away
```

`layer` is one of the eight `Layer` names, lowest first: `docDefaults`,
`tableStyle`, `paragraphStyle`, `styleNumbering`, `numbering`,
`directParagraph`, `runStyle`, `directRun`.

The target can be a `Paragraph`, a `Run`, or a table `_Cell`. For a `Run`
you get run-direct formatting included; for a `Paragraph` it stops at
paragraph level.

### Is this direct formatting doing anything?

`stop_below` resolves the target as if that layer and everything above it
were absent — the "what would this look like without its own override?"
question:

```python
full = resolve_effective_formatting(run)
base = resolve_effective_formatting(run, stop_below="directRun")

if full.font_size == base.font_size:
    print("the direct size is redundant — the style already says that")
```

This is the primitive the [linter](linting.md) is built on.

### Spacing needs a second call

`spacing_before` / `spacing_after` are what the cascade *declares*. Whether
either is actually applied depends on the paragraph's neighbours, so ask
separately:

```python
from docx_plus.styles import resolve_paragraph_spacing

s = resolve_paragraph_spacing(p)
s.space_above, s.space_below       # twips actually applied, both edges
s.declared_before, s.declared_after
s.contextual_spacing
s.before_suppressed, s.after_suppressed
```

!!! warning "Do not add space-after to the next paragraph's space-before"
    Word **tops the first up to the second** rather than summing them, so a
    pair of paragraphs sits `max(after, before)` apart. And
    `<w:contextualSpacing>` drops a paragraph's space on any edge where the
    neighbour has the same `styleId` — which affects any document using
    Word's built-in list styles, since fourteen of them carry the flag.

    Both rules were measured against live Word. See [the concept
    page](../concepts/cascade.md#paragraph-spacing-the-one-property-the-cascade-cannot-finish).

## Sweeping a whole document

`resolve_effective_formatting` rebuilds every document-level lookup on each
call — the theme, the styles part, each `basedOn` chain. That is right for
one paragraph and wrong for all of them. `iter_resolved_paragraphs`
resolves the whole document against one shared cache:

```python
from docx_plus.styles import iter_resolved_paragraphs

for resolved in iter_resolved_paragraphs(doc, include_baseline=True):
    print(resolved.index, resolved.formatting.style_id, resolved.text)

    for run in resolved.runs:
        if run.baseline and run.formatting.bold != run.baseline.bold:
            print("   run", run.index, "sets bold directly")
```

`include_baseline=True` resolves each target a second time with its own
direct layer excluded, which is what makes "is this override doing
anything?" answerable across a whole document. It roughly doubles the work,
so it is off by default.

!!! note "`index` is not an index into `doc.paragraphs`"
    The sweep descends into table cells, which `doc.paragraphs` omits, so
    the two diverge at the first table. Pass `include_tables=False` to make
    them line up.

    The sweep is **body only** — headers, footers, footnotes, endnotes, and
    comments live in separate parts and are not visited.

`resolved.runs` includes runs inside a `<w:hyperlink>`, which
`Paragraph.runs` omits; together they cover exactly the text `.text`
reports.

## Creating and applying styles

The Word-native workflow: define a style once, apply it, and later change
the *style* to restyle every paragraph that uses it.

```python
from docx import Document
from docx_plus.styles import apply_style, create_style

doc = Document()
create_style(
    doc, "BrandHeading",
    style_type="paragraph",        # required: "paragraph" | "character" | "table"
    based_on="Heading1",
    name="Brand Heading",          # the UI name; defaults to the style_id
    font_name="Inter",
    font_size=18.0,
    color_rgb="2F5496",
    bold=True,
    spacing_after=240,
)
apply_style(doc.add_paragraph("Hello"), "BrandHeading")
doc.save("out.docx")
```

- `create_style(doc, style_id, *, style_type, name=None, based_on=None,
  next_style=None, linked_style=None, ui_priority=None, q_format=None,
  custom=None, **properties)` — raises `StyleExistsError` if the id is
  taken.
- `modify_style(doc, style_id, *, if_missing="raise", **properties)` —
  `w:ind` / `w:spacing` / `w:rFonts` merge rather than replace. Pass
  `if_missing="create"` to define the style instead of raising.
- `apply_style(target, style_id)` — a `Paragraph`, `Run`, or `_Cell`.
- `delete_style(doc, style_id, *, force=False)` — raises `StyleInUseError`
  if the style is still referenced.

### Which properties are accepted

Field names match `ResolvedFormatting`, so resolved output round-trips
straight back into the modifier.

- **Paragraph:** `alignment`, `indent_left`, `indent_right`,
  `indent_first_line`, `spacing_before`, `spacing_after`, `line_spacing`,
  `line_spacing_rule`, `contextual_spacing`, `keep_with_next`,
  `keep_lines`, `page_break_before`, `outline_level`
- **Run:** `font_name`, `font_size`, `bold`, `italic`, `underline`,
  `strike`, `color_rgb`, `highlight`, `caps`, `small_caps`, `vanish`,
  `vert_align`

The decorative and complex-script toggles (`cs_bold`, `cs_italic`,
`emboss`, `imprint`, `outline`, `shadow`) are **read-only** — reported by
`resolve_effective_formatting` but not accepted as kwargs here.

### Restyling an existing document

Change the style, not the paragraphs:

```python
from docx import Document
from docx_plus.styles import ensure_style, modify_style

doc = Document("report.docx")
ensure_style(doc, "Heading1")
modify_style(doc, "Heading1", color_rgb="C00000", font_size=20.0,
             spacing_before=480)
doc.save("restyled.docx")   # every Heading1 paragraph re-renders on open
```

## Materialising latent built-in styles

Word's built-ins (`Heading1`–`Heading9`, `Title`, `Quote`, `TOC1`–`TOC9`,
`Caption`, `FootnoteText`, …) are *latent*: defined by Word's defaults but
absent from `styles.xml` until something uses them. `ensure_style` knows
**107** of them with real Word-extracted defaults, and is idempotent:

```python
from docx_plus.styles import ensure_style

ensure_style(doc, "Heading1")     # materialise if absent; no-op if present
ensure_style(doc, "TOC2")
ensure_style(doc, "Caption")

ensure_style(doc, "Quote", color_rgb="404040")   # override, only when creating
```

!!! note "`ensure_style` never overwrites"
    If the style is already defined — including the Word-2007-vintage
    definitions python-docx's own template ships — the existing definition
    is returned unchanged. It is a "the style is absent, here is what Word
    would have written" fallback, not a "force my defaults" mechanism. For
    that, use `modify_style`.

The full tiered table is on the [styles concept
page](../concepts/styles.md#the-built-in-styles-table).

## Working with someone else's document

The same logical style shows up as `Heading1` in one document, `Heading 1`
in another, and `HeadingOne` in a third. Code that calls
`apply_style(p, "Heading1")` fails against the second — not because the
style is missing, but because the id doesn't match.

```python
from docx_plus.styles import find_matching_style, list_styles, remap_styles

# Case- and space-insensitive lookup. Safe to call unconditionally.
proxy = find_matching_style(doc, "Heading1")   # StyleProxy | None

# Bulk-normalise the document's body references onto canonical ids:
remap_styles(doc, mapping={"Heading 1": "Heading1"})

for s in list_styles(doc, style_type="paragraph", include_latent=True):
    print(s.style_id, s.name, s.is_latent)
```

`remap_styles(doc, *, targets=None, mapping=None, create_missing=False)`
rewrites **body references only** (`w:pStyle`, `w:rStyle`, `w:tblStyle`).
Style-to-style references inside `styles.xml` are deliberately left alone —
see [why](../concepts/styles.md#style-remapping).

The same job from a shell is [`docx-plus restyle`](../cli.md).

## Theme colours (read-only)

```python
from docx_plus.styles.theme import load_theme, resolve_theme_color, resolve_theme_font

theme = load_theme(doc)                 # None if the part is missing or malformed
if theme:
    accent = resolve_theme_color(theme, "accent1", tint="99")   # -> "RRGGBB"
    body = resolve_theme_font(theme, "minorHAnsi")              # e.g. "Calibri"
```

`tint` and `shade` are **hex byte strings** `"00"`–`"FF"`, not ints: `tint`
lightens toward white (`"FF"` is a no-op), `shade` darkens toward black.

Name-to-slot resolution is per-document, not a fixed alias table — a
document can remap `text1` to white via `settings.xml`. Theme *writing* is
out of scope.

## Errors

All subclass `DocxPlusError`: `StyleExistsError`, `StyleNotFoundError`,
`StyleInUseError`, `StyleCascadeError` (a `basedOn` cycle, or a chain
deeper than 11), `UnknownStylePropertyError` (also `TypeError`),
`InvalidColorError` (also `ValueError`), `ThemeError`.

## See also

- [How the cascade works](../concepts/cascade.md) — the six layers, the
  toggle rule, spacing arithmetic, provenance
- [Remapping and built-ins](../concepts/styles.md) — the design behind
  `remap_styles` and the 107-style table
- Reference: [`styles.inspect`](../reference/styles-inspect.md),
  [`styles.modify`](../reference/styles-modify.md),
  [`styles.sweep`](../reference/styles-sweep.md),
  [`styles.theme`](../reference/styles-theme.md)
- Examples: `inspect_document.py`, `restyle_existing.py`
