# docx_plus

OOXML-level extensions for [python-docx](https://python-docx.readthedocs.io/).
Composes with python-docx rather than replacing it: callers keep their
`Document` object and use `docx_plus` for the operations python-docx
can't reach.

**v0.1 capabilities** (in progress — see roadmap below):

- **Style cascade**: read the effective formatting that would apply to
  any paragraph/run/cell, with per-field provenance; modify styles in
  the Word-native way rather than scattering direct formatting.
- **Content controls** (Phase 4, not yet shipped): create text /
  dropdown / date / checkbox controls; read values back; enforce form
  protection.
- **Fields** (Phase 5, not yet shipped): insert simple fields,
  mark them dirty so Word recalculates them on next open.

> **Status:** early development (v0.1 in progress). Pre-publication —
> not yet on PyPI. Read [`SPEC.md`](SPEC.md) for the API contract and
> [`IMPLEMENTATION.md`](IMPLEMENTATION.md) for the build plan.

## Install (development)

```bash
git clone <repo-url> docx_plus
cd docx_plus
uv sync --extra dev      # or: pip install -e ".[dev]"
```

## 60-second quickstart

### Inspect: why does this paragraph look the way it does?

```python
from docx import Document
from docx_plus.styles import resolve_effective_formatting

doc = Document("report.docx")
p = doc.paragraphs[0]

resolved = resolve_effective_formatting(p, include_provenance=True)
print(resolved.style_name)              # e.g. "Title"
print(resolved.font_size)               # e.g. 28.0  (points)
print(resolved.bold)                    # True / False / None
print(resolved.provenance["font_size"]) # FormattingSource(layer='paragraphStyle', ...)
```

`ResolvedFormatting` carries every formatting field that the OOXML
cascade can set — `font_name`, `font_size`, `bold`, `italic`, `color_rgb`,
`alignment`, `indent_*`, `spacing_*`, `line_spacing`, plus run-level
toggles. With `include_provenance=True`, every populated field is
keyed in `.provenance` to the cascade layer (and style ID) that
contributed it. That's how you answer "why is this paragraph 14pt
italic?" — the provenance tells you exactly which style in the
basedOn chain set the size and whether the italic came through XOR.

### Modify: define a custom heading and apply it

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
    spacing_after=240,
)

p = doc.add_paragraph("Hello, world")
apply_style(p, "BrandHeading")
doc.save("out.docx")
```

This is the Word-native workflow: define a style, apply it. Changing
the style later changes every paragraph that uses it, not just the
ones you remember to update.

### Ensure: materialise a built-in latent style

Word's built-ins (`Heading1`–`Heading9`, `Title`, `Quote`, `TOC1`–`TOC9`,
`FootnoteText`, `BlockText`, `PlainText`, …) are *latent* — defined by
Word's defaults but not actually present in `styles.xml` until they're
used. `ensure_style` knows about **107** of them, with defaults
extracted from real Word-saved samples (not guessed):

```python
from docx import Document
from docx_plus.styles import ensure_style, apply_style

doc = Document()
ensure_style(doc, "Heading1")           # idempotent — materialises if absent
ensure_style(doc, "Heading1")           # ...no-op the second time
ensure_style(doc, "TOC2")               # also works for less-common built-ins
ensure_style(doc, "BlockText")
apply_style(doc.add_paragraph("Intro"), "Heading1")
```

The full list is tiered in [`docs/ARCHITECTURE.md` §5](docs/ARCHITECTURE.md#5-built-in-styles-table)
— Core/A–G cover essentially every style a Word user reaches for.

For documents authored elsewhere where IDs may not match (e.g. style
named `"Heading 1"` with a space), `ensure_style(doc, "Heading1",
match_existing=True)` will find the existing definition via case- and
space-insensitive matching, or use [`remap_styles`](docs/ARCHITECTURE.md#4-style-remapping-phase-35)
for document-wide normalisation.

## Roadmap

| Phase | Capability | Status |
|---|---|---|
| 1 | Foundation (`core/ns`, `core/oxml`, `core/ids`, `_testing/`) | ✓ complete |
| 2 | Style inspection (`styles/inspect`, `styles/theme`) | ✓ complete |
| 3 | Style modification (`styles/modify`) | ✓ complete |
| 3.5 | Style remapping (`find_matching_style`, `remap_styles`, `ensure_style(match_existing=)`) | ✓ complete |
| 4 | Content controls (`controls/`) | ✓ complete |
| 5 | Fields + document protection (`fields/`, `protection/`) | not started |
| 6 | Polish — examples, headless LibreOffice smoke tests, CI doc build | not started |

## Documentation

- [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) — module layout,
  cascade algorithm, schema-strict insertion, error hierarchy,
  invariants
- [`docs/API.md`](docs/API.md) — hand-curated index of every public
  symbol with links to the auto-generated reference
- [`docs/TEST_GAPS.md`](docs/TEST_GAPS.md) — honest accounting of
  where the test suite has real holes (snapshot at end of Phase 3.5)
- `docs/reference/` — per-module API reference, rendered by
  [MkDocs](https://www.mkdocs.org) + [mkdocstrings](https://mkdocstrings.github.io).
  `uv run mkdocs serve` to browse locally.

## License

MIT (placeholder — confirm before publishing). See [`LICENSE`](LICENSE).
