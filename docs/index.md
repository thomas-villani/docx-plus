# docx_plus

OOXML-level extensions for [python-docx](https://python-docx.readthedocs.io/) —
the library every python-docx power user ends up writing badly: hardened
helpers for OOXML operations that sit just past python-docx's
abstraction boundary.

Status: **v0.1 complete**. Pre-publication — not yet on PyPI.

## What's in v0.1

- **Style cascade** — read the effective formatting that would apply to
  any paragraph/run/cell, with per-field provenance; modify styles in
  the Word-native way rather than scattering direct formatting.
- **Content controls** — text / dropdown / date / checkbox controls
  via `FormBuilder`; round-trip read/write of values; form protection.
- **Fields** — PAGE / NUMPAGES / DATE / generic complex fields; mark
  fields dirty so Word recalculates them on next open.
- **Protection** — enforce form-fill, read-only, comments-only, or
  tracked-changes mode at the document level.

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
