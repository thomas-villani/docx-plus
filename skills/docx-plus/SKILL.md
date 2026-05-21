---
name: docx-plus
description: >-
  Use when generating or editing Microsoft Word (.docx) files in Python and you
  need OOXML features beyond what python-docx exposes: footnotes, endnotes,
  table of contents, captions / table of figures, fillable forms (content
  controls: text / dropdown / date / checkbox), anchored comments, style-cascade
  inspection and Word-native style creation / modification, bookmarks and
  cross-references, multi-column layout, page borders, line numbering,
  page-number and date fields, and document protection. docx_plus composes with
  python-docx — you keep your own Document object and call docx_plus for the
  operations python-docx can't reach.
---

# docx_plus

`docx_plus` is an OOXML-level extension layer for
[python-docx](https://python-docx.readthedocs.io/). It does **not** replace
python-docx: you create and own a `docx.Document`, add ordinary content
(headings, paragraphs, tables, runs) with python-docx as usual, and call
`docx_plus` functions for the things python-docx can't do. Every function takes
a python-docx object (a `Document`, `Paragraph`, `Run`, `Section`, or `_Cell`)
and mutates it in place.

```python
from docx import Document
from docx_plus.notes import add_footnote

doc = Document()                          # python-docx, as normal
p = doc.add_paragraph("A claim")          # python-docx, as normal
add_footnote(p, "Source: internal data") # docx_plus does what python-docx can't
doc.save("out.docx")                      # python-docx, as normal
```

## Install

```bash
uv add docx-plus          # or: pip install docx-plus
```

Import name is `docx_plus` (underscore); distribution name is `docx-plus`
(hyphen). It depends on `python-docx`, which it imports as `docx`.

## Conventions that apply everywhere — read these first

These trip up every first attempt. They hold across all modules.

- **Composition, not subclassing.** There is no `DocxPlus` document class. Keep
  your `docx.Document`. `FormBuilder` is the one wrapper, and it exposes the
  underlying document as `fb.doc`.
- **Fields don't show a value until Word recalculates them.** Anything that
  inserts a *field* — `add_toc`, `add_caption`, `add_table_of_figures`,
  `add_cross_reference`, `add_page_number_field`, `add_date_field`, `add_field`
  — produces an empty/placeholder result on disk. Call
  `docx_plus.fields.mark_fields_dirty(doc)` once, after all inserts and before
  `save`, so Word fills them in on open (or the user presses F9). Forget this
  and the TOC / page numbers / cross-refs look blank.
- **Units.** Font size is **points** (`float`). Spacing, indents, column gaps,
  border offsets, and line-number distance are **twips** (`int`; 1440 twips =
  1 inch, 20 twips = 1 pt). Border thickness is **eighths of a point** (`int`).
  Colors are uppercase `"RRGGBB"` hex strings with **no** leading `#`.
- **Style IDs vs. style names.** Functions take the machine-readable
  `w:styleId` (e.g. `"Heading1"`, no space), *not* the UI name (`"Heading 1"`).
  For documents authored elsewhere, reconcile with
  `ensure_style(..., match_existing=True)` or `remap_styles` (see
  `reference/styles.md`).
- **Toggle properties** (`bold`, `italic`, `caps`, `small_caps`, `strike`,
  `vanish`, …): in `create_style` / `modify_style`, `True` sets it, `False`
  forces it off, and `None` (in `modify_style`) *removes* the element so it
  inherits from the parent style again.
- **Sharing an id registry for batch inserts.** Comments, bookmarks, and notes
  each take an optional `id_registry=`. When you add several in one session,
  build one registry and pass it to every call so the allocated `w:id`s stay
  unique. The per-module reference files show this.
- **Target is Word.** OOXML is rendered by Microsoft Word. LibreOffice / Pages
  mostly work but aren't the contract.
- **Errors.** Every typed error subclasses `docx_plus.DocxPlusError`; most also
  subclass the matching builtin (`ValueError`, `KeyError`, `TypeError`) so
  ordinary `except` clauses still catch them.

## Capability map

Load the reference file for the area you're working in — each is a complete,
copy-pasteable guide to that module's public API.

| If you need to…                                              | Use module        | Reference file              |
| ------------------------------------------------------------ | ----------------- | --------------------------- |
| Build a fillable form; read / set / clear control values; lock a document | `controls`, `protection` | `reference/forms.md`     |
| Inspect why a paragraph looks the way it does; create / modify / apply styles; resolve theme colors | `styles`         | `reference/styles.md`       |
| Table of contents, figure/table captions, table of figures, footnotes, endnotes, bookmarks, cross-references, page-number/date/generic fields | `publishing`, `notes`, `bookmarks`, `fields` | `reference/publishing.md`   |
| Multi-column sections, mid-document section breaks, distinct even/odd headers, line numbering, page borders | `layout`         | `reference/layout.md`       |
| Anchored review comments that "show in document" correctly   | `comments`        | `reference/comments.md`     |

## Two patterns worth memorizing

**Inserting fields → always mark dirty before save:**

```python
from docx import Document
from docx_plus.fields import mark_fields_dirty
from docx_plus.publishing import add_toc

doc = Document()
doc.add_heading("Contents", level=1)
add_toc(doc.add_paragraph(), levels=(1, 2))
# ... add headings the TOC will collect ...
mark_fields_dirty(doc)          # <-- without this the TOC renders blank
doc.save("report.docx")
```

**The Word-native styling workflow → change the style, not each paragraph:**

```python
from docx import Document
from docx_plus.styles import ensure_style, modify_style, apply_style

doc = Document()
ensure_style(doc, "Heading1")                 # materialize the latent built-in
modify_style(doc, "Heading1", color_rgb="C00000", font_size=20.0, bold=True)
apply_style(doc.add_paragraph("Section"), "Heading1")
doc.save("out.docx")
```

## Reference files

- `reference/forms.md` — `FormBuilder` (text / dropdown / combobox / date /
  checkbox), `read_controls` / `set_control_value` / `clear_control`, and
  document protection.
- `reference/styles.md` — `resolve_effective_formatting` with provenance, plus
  `create_style` / `modify_style` / `apply_style` / `ensure_style` /
  `remap_styles` and read-only theme resolution.
- `reference/publishing.md` — TOC, captions, table of figures, footnotes,
  endnotes, bookmarks, cross-references, and the field helpers
  (`mark_fields_dirty` lives here).
- `reference/layout.md` — columns, section breaks, even/odd headers, line
  numbering, page borders.
- `reference/comments.md` — anchored comments (add / edit / delete / read).

For exhaustive signatures and the error taxonomy, the rendered docs are at
<https://thomas-villani.github.io/docx-plus/> and the in-repo index is
`docs/API.md`. Runnable end-to-end examples live in `docx_plus/examples/`.
