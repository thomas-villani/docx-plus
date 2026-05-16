# docx_plus

OOXML-level extensions for [python-docx](https://python-docx.readthedocs.io/) —
the library every python-docx power user ends up writing badly: hardened
helpers for OOXML operations that sit just past python-docx's
abstraction boundary.

Status: **early development** (v0.1 in progress). Pre-publication — not
yet on PyPI.

## What's in v0.1

- **Style cascade** — read the effective formatting that would apply to
  any paragraph/run/cell, with per-field provenance; modify styles in
  the Word-native way rather than scattering direct formatting.
- **Content controls** *(Phase 4, not yet shipped)* — text / dropdown /
  date / checkbox controls; read values back; form protection.
- **Fields** *(Phase 5, not yet shipped)* — page numbers, dates,
  generic fields; mark fields dirty so Word recalculates on next open.

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

ensure_style(doc, "Heading1")  # materialises Word's built-in if missing
```

## Roadmap

| Phase | Capability | Status |
|---|---|---|
| 1 | Foundation | ✓ complete |
| 2 | Style inspection | ✓ complete |
| 3 | Style modification | ✓ complete |
| 3.5 | Style remapping | ✓ complete |
| 4 | Content controls | not started |
| 5 | Fields + protection | not started |
| 6 | Polish (examples, smoke tests, CI doc build) | not started |
