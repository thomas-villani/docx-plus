# Styles — cascade inspection, modification, theme

Module: `docx_plus.styles`. Two halves: **inspection** (read the *effective*
formatting that the OOXML cascade would apply, with per-field provenance) and
**modification** (create / modify / apply / remove styles the Word-native way).
Plus read-only theme-color resolution.

## Inspecting effective formatting

`resolve_effective_formatting` walks the six cascade layers (document defaults →
table style → paragraph style chain via `basedOn` → numbering → paragraph
direct → run direct) and returns a single frozen `ResolvedFormatting` with the
value that actually wins for each field.

```python
from docx import Document
from docx_plus.styles import resolve_effective_formatting

doc = Document("report.docx")
p = doc.paragraphs[0]

r = resolve_effective_formatting(p)
print(r.style_id, r.style_name)   # "Title", "Title"
print(r.font_size)                # 28.0  (points)
print(r.bold)                     # True / False / None
```

`resolve_effective_formatting(target, *, include_provenance=False, table_context=None)`

- `target` — a `Paragraph`, a `Run`, or a table `_Cell`
  (`doc.tables[0].rows[0].cells[0]`). For a `Run`, run-direct formatting is
  included; for a `Paragraph`, it stops at paragraph level.
- `include_provenance=True` — also populate `.provenance`, a dict mapping each
  set field name to a `FormattingSource(layer, style_id, chain_depth,
  is_toggle_resolved)`. This is how you answer "*why* is this 14pt italic?" —
  the source names the cascade layer and style ID that set it.
- `table_context` — a `TableContext` to override the auto-derived cell position
  when resolving conditional table-style branches (`firstRow`, `lastRow`,
  banded fills, …).

```python
r = resolve_effective_formatting(p, include_provenance=True)
src = r.provenance["font_size"]
print(src.layer, src.style_id)    # e.g. "paragraphStyle", "Title"
```

### `ResolvedFormatting` fields

A value of `None` means "not set at any layer" (inherits Word's default).
Common fields:

- Identity: `style_id`, `style_name`
- Paragraph: `alignment`, `indent_left`, `indent_right`, `indent_first_line`,
  `spacing_before`, `spacing_after`, `line_spacing`, `line_spacing_rule`,
  `keep_with_next`, `keep_lines`, `page_break_before`, `outline_level`
- Run: `font_name`, `font_size`, `bold`, `italic`, `underline`, `strike`,
  `double_strike`, `color_rgb`, `highlight`, `vert_align`
- Toggles (all 12 ECMA-376 17.7.3): `bold`, `italic`, `cs_bold`, `cs_italic`,
  `caps`, `small_caps`, `strike`, `vanish`, `emboss`, `imprint`, `outline`,
  `shadow`
- Numbering: `num_id`, `num_level`
- Meta: `partial` (True if the theme part was missing/malformed),
  `provenance` (the dict above, or `None`)

Units: `font_size` in points; `indent_*` / `spacing_*` in twips;
`line_spacing` is twips unless `line_spacing_rule == "auto"`, where it's a
multiplier (e.g. `1.15`); `color_rgb` / `highlight` are `"RRGGBB"` hex.

## Creating, modifying, applying styles

The Word-native workflow: define a style once, apply it to paragraphs/runs, and
later change the *style* to restyle every paragraph that uses it.

```python
from docx import Document
from docx_plus.styles import create_style, apply_style

doc = Document()
create_style(
    doc, "BrandHeading",
    style_type="paragraph",     # required: "paragraph" | "character" | "table"
    based_on="Heading1",
    name="Brand Heading",       # UI name; defaults to the style_id
    font_name="Inter", font_size=18.0, color_rgb="2F5496",
    bold=True, spacing_after=240,
)
apply_style(doc.add_paragraph("Hello"), "BrandHeading")
```

`create_style(doc, style_id, *, style_type, name=None, based_on=None, next_style=None, linked_style=None, ui_priority=None, q_format=None, custom=None, **properties)` —
raises `StyleExistsError` if `style_id` already exists.

`modify_style(doc, style_id, *, if_missing="raise", **properties)` — change one
or more properties. `w:ind` / `w:spacing` / `w:rFonts` merge rather than
replace. `if_missing="create"` defines the style instead of raising
`StyleNotFoundError` when it doesn't exist yet (default is `"raise"`).

`apply_style(target, style_id)` — apply by ID to a `Paragraph`, `Run`, or
`_Cell`; raises `StyleNotFoundError` if undefined.

`delete_style(doc, style_id, *, force=False)` — raises `StyleInUseError` if the
style is referenced (unless `force=True`).

### Properties accepted by `create_style` / `modify_style`

Field names match `ResolvedFormatting`, so resolved output round-trips back into
the modifier.

- Paragraph: `alignment`, `indent_left`, `indent_right`, `indent_first_line`,
  `spacing_before`, `spacing_after`, `line_spacing`, `line_spacing_rule`,
  `keep_with_next`, `keep_lines`, `page_break_before`, `outline_level`
- Run: `font_name`, `font_size`, `bold`, `italic`, `underline`, `strike`,
  `color_rgb`, `highlight`, `caps`, `small_caps`, `vanish`, `vert_align`

The decorative/complex-script toggles (`cs_bold`, `cs_italic`, `emboss`,
`imprint`, `outline`, `shadow`) are **read-only** — surfaced by
`resolve_effective_formatting` but not accepted as kwargs here.

**Toggle semantics** (`bold`, `italic`, `caps`, `small_caps`, `strike`,
`vanish`): `True` sets it, `False` forces it off, and in `modify_style`, `None`
*removes* the element so it inherits from the parent style's value again.

### Restyling an existing document

```python
from docx import Document
from docx_plus.styles import ensure_style, modify_style

doc = Document("report.docx")
ensure_style(doc, "Heading1")
modify_style(doc, "Heading1", color_rgb="C00000", font_size=20.0,
             spacing_before=480)
doc.save("restyled.docx")   # every Heading1 paragraph re-renders on open
```

## Materializing latent built-in styles: `ensure_style`

Word's built-ins (`Heading1`–`Heading9`, `Title`, `Quote`, `TOC1`–`TOC9`,
`Caption`, `FootnoteText`, …) are *latent* — defined by Word's defaults but
absent from `styles.xml` until used. `ensure_style` knows 107 of them with
real Word-extracted defaults, and is idempotent.

```python
from docx_plus.styles import ensure_style

ensure_style(doc, "Heading1")                 # materialize if absent; no-op if present
ensure_style(doc, "TOC2")
ensure_style(doc, "Caption")
# Override defaults only when creating:
ensure_style(doc, "Quote", color_rgb="404040")
```

`ensure_style(doc, style_id, *, match_existing=False, **defaults_if_creating)`.
For documents authored elsewhere where the ID may differ (e.g. `"Heading 1"`
with a space), pass `match_existing=True` to find it via case- and
space-insensitive matching instead of creating a duplicate.

## Reconciling foreign style names

```python
from docx_plus.styles import find_matching_style, remap_styles, list_styles

find_matching_style(doc, "Heading1")   # case/space-insensitive lookup -> StyleProxy | None

# Bulk-normalize body references (e.g. "Heading 1" -> "Heading1"):
remap_styles(doc, mapping={"Heading 1": "Heading1"})

for s in list_styles(doc, style_type="paragraph", include_latent=True):
    print(s.style_id, s.name, s.is_latent)
```

- `list_styles(doc, *, style_type=None, include_latent=False) -> list[StyleInfo]`
  — `StyleInfo` has `style_id`, `name`, `style_type`, `based_on`, `is_default`,
  `is_latent`.
- `remap_styles(doc, *, targets=None, mapping=None, create_missing=False)` —
  bulk reconciliation; rewrites body references only.
- `StyleProxy` wraps a live `w:style` element; `.modify(**props)` and
  `.delete(force=...)` delegate to the module functions.

## Theme colors (read-only)

```python
from docx_plus.styles.theme import load_theme, resolve_theme_color, resolve_theme_font

theme = load_theme(doc)               # None if the part is missing/malformed
if theme:
    accent = resolve_theme_color(theme, "accent1", tint="99")   # -> "RRGGBB"
    body_font = resolve_theme_font(theme, "minorHAnsi")         # e.g. "Calibri"
```

`resolve_theme_color(theme, name, *, tint=None, shade=None)` — `name` is a Word
`ST_ThemeColor` (`"accent1"`, `"text1"`, …). `tint` / `shade` are **hex byte
strings** `"00"`–`"FF"` (not ints): `tint` lightens toward white (`"FF"` is a
no-op, `"00"` is pure white), `shade` darkens toward black. Returns `"RRGGBB"`
or `None`.

Also available: `ThemeColors.base(name)` / `.font(token)`, and the transforms
`apply_theme_tint`, `apply_theme_shade`, `apply_lum_mod`, `apply_lum_off`.
Theme *writing* is out of scope.

## Errors

All subclass `DocxPlusError`. `StyleExistsError`, `StyleNotFoundError`,
`StyleInUseError`, `StyleCascadeError` (basedOn cycle or chain depth > 11),
`UnknownStylePropertyError` (`TypeError`; unknown `**properties` kwarg),
`InvalidColorError` (`ValueError`; not valid `RRGGBB`), `ThemeError`.

See also: `docx_plus/examples/inspect_document.py` and `restyle_existing.py`.
