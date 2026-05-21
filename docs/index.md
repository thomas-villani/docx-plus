# docx_plus

OOXML-level extensions for [python-docx](https://python-docx.readthedocs.io/) —
the library every python-docx power user ends up writing badly: hardened
helpers for OOXML operations that sit just past python-docx's
abstraction boundary.

Status: **v0.2.0 released** — published on
[PyPI](https://pypi.org/project/docx-plus/).

## Capabilities

- **Style cascade** — read the effective formatting that would apply to
  any paragraph/run/cell, with per-field provenance; modify styles in
  the Word-native way rather than scattering direct formatting.
- **Content controls** — text / dropdown / date / checkbox controls
  via `FormBuilder`; round-trip read/write of values; form protection.
- **Fields** — PAGE / NUMPAGES / DATE / generic complex fields; mark
  fields dirty so Word recalculates them on next open.
- **Protection** — enforce form-fill, read-only, comments-only, or
  tracked-changes mode at the document level.
- **Anchored comments** (v0.2) — the body-side range markers
  python-docx skips, so "show in document" actually works.
- **Layout** (v0.2) — multi-column sections, mid-document section
  breaks, distinct even/odd headers.
- **Bookmarks + cross-references** (v0.2) — paired body markers plus
  `REF` / `PAGEREF` fields.
- **Footnotes + endnotes** (v0.2) — separate `footnotes.xml` /
  `endnotes.xml` parts; insert + edit in-place.
- **Layout: line numbers + page borders** (v0.2) — marginal line
  numbering and decorative page borders.
- **Conditional table-style formatting** (v0.2) — the cascade
  applies `<w:tblStylePr>` branches (first row, banded rows, corners)
  in ECMA-376 17.7.6.5 precedence order.
- **Publishing** (v0.2) — Table of Contents, figure / table captions
  via `SEQ`, and a downstream Table of Figures.

## Where to start

- New to the library? Read the **[Architecture](ARCHITECTURE.md)**
  overview and skim the API index.
- Want the full reference? **[API Index](API.md)** lists every public
  symbol; **[Reference](reference/core-ns.md)** has per-module pages
  with full signatures and docstrings.
- Auditing the project? See the **[Test Gaps](TEST_GAPS.md)** snapshot.

## Quickstart

### Inspect

```python
from docx import Document
from docx_plus.styles import resolve_effective_formatting

doc = Document("report.docx")
p = doc.paragraphs[0]
resolved = resolve_effective_formatting(p, include_provenance=True)

print(resolved.font_size, resolved.bold)
print(resolved.provenance["font_size"])  # which cascade layer set it
```

### Modify

```python
from docx import Document
from docx_plus.styles import create_style, apply_style

doc = Document()
create_style(
    doc, "BrandHeading",
    style_type="paragraph",
    based_on="Heading1",
    font_name="Inter",
    font_size=18.0,
    color_rgb="2F5496",
    bold=True,
)
apply_style(doc.add_paragraph("Hello"), "BrandHeading")
doc.save("out.docx")
```

### Ensure (latent built-ins)

```python
from docx_plus.styles import ensure_style

ensure_style(doc, "Heading1")  # one of 107 known built-ins
ensure_style(doc, "TOC2")
ensure_style(doc, "BlockText")
```

See [`ARCHITECTURE.md` §5](ARCHITECTURE.md#5-built-in-styles-table) for
the full tiered table.

### Forms

```python
from docx_plus.controls import FormBuilder

fb = FormBuilder()
p = fb.doc.add_paragraph("Name: ")
fb.add_text_control(p, tag="name", placeholder="Type your name")
p = fb.doc.add_paragraph("Dept: ")
fb.add_dropdown(p, tag="dept", items=["Eng", "Design", "Ops"])
fb.save("form.docx")
```

Read / update with `read_controls(doc)` and `set_control_value(doc,
tag, value)`. See [`ARCHITECTURE.md` §6](ARCHITECTURE.md#6-content-controls).

### Fields & protection

```python
from docx_plus.fields import add_page_number_field, mark_fields_dirty
from docx_plus.protection import protect_document

p = doc.add_paragraph("Page ")
add_page_number_field(p)
p.add_run(" of ")
add_page_number_field(p, field="NUMPAGES")

mark_fields_dirty(doc)               # Word recalculates on open
protect_document(doc, mode="forms")  # only content controls editable
```

See [`ARCHITECTURE.md` §7](ARCHITECTURE.md#7-fields-and-protection).

### Comments (v0.2)

```python
from docx import Document
from docx_plus.comments import add_comment, read_comments

doc = Document()
p = doc.add_paragraph()
p.add_run("Project Apollo ")
target = p.add_run("ships next quarter")
add_comment(target, "Optimistic — let's see QA.", author="Alice")

for c in read_comments(doc):
    print(c.author, c.text, "→", c.anchored_text)
```

Unlike python-docx's own `add_comment` (which only writes the part-side
body), `docx_plus` writes the three body-side anchors — so Word's
"show in document" jumps to the right place. See
[`ARCHITECTURE.md` §7.6](ARCHITECTURE.md#76-anchored-comments).

### Layout (v0.2)

```python
from docx_plus.layout import insert_section_break, set_columns

split = doc.add_paragraph("Section break here ↓")
new_section = insert_section_break(split, start_type="continuous")
set_columns(new_section, 2, space=720, separator=True)
```

See [`ARCHITECTURE.md` §7.7](ARCHITECTURE.md#77-layout).

### Bookmarks + cross-references (v0.2)

```python
from docx_plus.bookmarks import add_bookmark, add_cross_reference
from docx_plus.fields import mark_fields_dirty

heading = doc.add_heading("Introduction", level=1)
add_bookmark(heading, "intro_section")

p = doc.add_paragraph("See ")
add_cross_reference(p, bookmark="intro_section", kind="text")
mark_fields_dirty(doc)   # Word recalculates REF / PAGEREF on open
```

See [`ARCHITECTURE.md` §7.8](ARCHITECTURE.md#78-bookmarks-and-cross-references).

### Footnotes + endnotes (v0.2)

```python
from docx_plus.notes import add_footnote, add_endnote, edit_footnote

p = doc.add_paragraph("This claim has a footnote")
ref = add_footnote(p, "Sourced from internal benchmarks, 2026-05-19.")
add_endnote(p, "Re-validated against external dataset Q3 2026.")

# Need to update the footnote body later?
edit_footnote(doc, ref.note_id, "Re-sourced from external benchmarks.")
```

See [`ARCHITECTURE.md` §7.9](ARCHITECTURE.md#79-footnotes-and-endnotes).

### Layout: line numbers + page borders (v0.2)

```python
from docx_plus.layout import Border, set_line_numbering, set_page_borders

set_line_numbering(doc.sections[0], count_by=5, restart="newPage")

rule = Border(style="single", size=8, color="2F5496")
set_page_borders(
    doc.sections[0], top=rule, bottom=rule, left=rule, right=rule,
)
```

See [`ARCHITECTURE.md` §7.7](ARCHITECTURE.md#77-layout).

### Publishing (v0.2)

```python
from docx_plus.fields import mark_fields_dirty
from docx_plus.publishing import add_caption, add_table_of_figures, add_toc

doc.add_heading("Contents", level=1)
add_toc(doc.add_paragraph(), levels=(1, 2))

doc.add_heading("Body", level=1)
cap = doc.add_paragraph()
add_caption(cap, "Figure ", caption_type="Figure")
cap.add_run(": System overview.")

doc.add_heading("List of Figures", level=1)
add_table_of_figures(doc.add_paragraph())

mark_fields_dirty(doc)   # Word populates TOC / SEQ / ToF on open
```

See [`ARCHITECTURE.md` §7.10](ARCHITECTURE.md#710-publishing).

## Roadmap

| Phase | Capability | Status |
|---|---|---|
| 1 | Foundation | ✓ complete |
| 2 | Style inspection | ✓ complete |
| 3 | Style modification | ✓ complete |
| 3.5 | Style remapping | ✓ complete |
| 4 | Content controls | ✓ complete |
| 5 | Fields + protection | ✓ complete |
| 6 | Polish (examples, smoke tests, CI doc build) | ✓ complete |
| v0.2 | Comments, layout, bookmarks/cross-refs, notes, `core/parts` | ✓ complete |
| v0.2 expansion | Toggle props, in-place edits, line numbering, page borders, conditional table styles, `publishing/` | ✓ complete |
