# docx_plus

OOXML-level extensions for [python-docx](https://python-docx.readthedocs.io/).
Composes with python-docx rather than replacing it: callers keep their
`Document` object and use `docx_plus` for the operations python-docx
can't reach.

**Capabilities** (v0.1 through v0.3):

- **Style cascade**: read the effective formatting that would apply to
  any paragraph/run/cell, with per-field provenance; modify styles in
  the Word-native way rather than scattering direct formatting.
- **Content controls**: build text / dropdown / date / checkbox
  controls with `FormBuilder`; read their values back; round-trip them
  through save/reopen.
- **Fields**: insert PAGE / NUMPAGES / DATE / generic complex fields;
  mark fields dirty so Word recalculates them on next open.
- **Protection**: enforce form-fill, read-only, comments-only, or
  tracked-changes mode at the document level.
- **Anchored comments** (v0.2): the body-side range markers
  python-docx skips, so "show in document" actually works.
- **Tracked changes** (v0.3): mark runs as insertions / deletions, read
  every revision with its author / timestamp / text, and accept or reject
  them — the inline `w:ins` / `w:del` revision marks python-docx can't reach.
- **Layout**: multi-column sections, mid-document section breaks,
  distinct even/odd headers (v0.2).
- **Bookmarks + cross-references**: paired body markers plus
  `REF` / `PAGEREF` fields (v0.2).
- **Footnotes + endnotes**: insert-only API backed by the separate
  ``footnotes.xml`` / ``endnotes.xml`` parts; in-place edits via
  `edit_footnote` / `edit_endnote` (v0.2).
- **Layout extras** (continued): line numbering (`set_line_numbering`)
  and page borders (`set_page_borders` + `Border` dataclass) (v0.2).
- **Conditional table-style formatting**: the cascade resolver applies
  `<w:tblStylePr>` branches (`firstRow`, `lastRow`, banded fills,
  corners) in ECMA-376 17.7.6.5 precedence order (v0.2).
- **Publishing primitives**: Table of Contents (`add_toc`), figure /
  table captions (`add_caption`), Table of Figures
  (`add_table_of_figures`) (v0.2).
- **Command line** (v0.3): a `docx-plus` console command over the
  library — `inspect` (effective formatting), `restyle` (style
  remapping), and `controls` (list / set / clear control values).

> **Status:** v0.3.0 is the current release, published on 2026-06-15 to
> [PyPI](https://pypi.org/project/docx-plus/). Read [`SPEC.md`](SPEC.md) for
> the API contract and [`IMPLEMENTATION.md`](IMPLEMENTATION.md) for the
> build plan.

## Install (development)

```bash
git clone https://github.com/thomas-villani/docx-plus.git
cd docx-plus
uv sync --extra dev          # or: pip install -e ".[dev]"
uv run pre-commit install    # run ruff check + ruff format on every commit
```

The pre-commit hooks mirror the CI lint gate (`ruff check` and
`ruff format`), so formatting issues are caught locally instead of on CI.
Run them against the whole tree any time with
`uv run pre-commit run --all-files`.

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

The full list is tiered in [Architecture §5](https://thomas-villani.github.io/docx-plus/ARCHITECTURE/#5-built-in-styles-table)
— Core/A–G cover essentially every style a Word user reaches for.

For documents authored elsewhere where IDs may not match (e.g. style
named `"Heading 1"` with a space), `ensure_style(doc, "Heading1",
match_existing=True)` will find the existing definition via case- and
space-insensitive matching, or use [`remap_styles`](https://thomas-villani.github.io/docx-plus/ARCHITECTURE/#4-style-remapping-phase-35)
for document-wide normalisation.

### Forms: build a fillable document with `FormBuilder`

```python
from docx_plus.controls import FormBuilder

fb = FormBuilder()  # or FormBuilder("template.docx")
fb.doc.add_heading("New employee form", level=1)

p = fb.doc.add_paragraph("Full name: ")
fb.add_text_control(p, tag="full_name", placeholder="Type your name")

p = fb.doc.add_paragraph("Department: ")
fb.add_dropdown(p, tag="dept", items=["Engineering", "Design", "Ops"])

p = fb.doc.add_paragraph("Start date: ")
fb.add_date_picker(p, tag="start_date", date_format="M/d/yyyy")

p = fb.doc.add_paragraph("Remote? ")
fb.add_checkbox(p, tag="remote", checked=False)

fb.save("form.docx")
```

Read or update an existing form's values with `read_controls` /
`set_control_value`:

```python
from docx import Document
from docx_plus.controls import read_controls, set_control_value

doc = Document("form.docx")
set_control_value(doc, "full_name", "Ada Lovelace")
set_control_value(doc, "dept", "Engineering")
doc.save("form_filled.docx")

values = read_controls(Document("form_filled.docx"))
print(values["full_name"].value)   # 'Ada Lovelace'
print(values["dept"].value)        # 'Engineering'
```

### Fields and protection: page numbers + lock-down

```python
from docx import Document
from docx_plus.fields import add_page_number_field, mark_fields_dirty
from docx_plus.protection import protect_document

doc = Document()
p = doc.add_paragraph("Page ")
add_page_number_field(p)
p.add_run(" of ")
add_page_number_field(p, field="NUMPAGES")

mark_fields_dirty(doc)               # Word recalculates fields on open
protect_document(doc, mode="forms")  # only content controls editable

doc.save("report.docx")
```

`add_date_field` and the generic `add_field(instruction=..., initial_text=...)`
cover dates and any other complex field (TOC, REF, MERGEFIELD, …).
`unprotect_document(doc)` removes any protection;
`is_protected(doc)` is a one-liner predicate.

### Comments: anchor reviewer feedback to specific runs

```python
from docx import Document
from docx_plus.comments import add_comment, read_comments

doc = Document()
p = doc.add_paragraph()
p.add_run("Project Apollo ")
target = p.add_run("ships next quarter")
p.add_run(".")

add_comment(target, "Optimistic — let's see what QA says.", author="Alice")
doc.save("review.docx")

for c in read_comments(Document("review.docx")):
    print(f"{c.author}: {c.text!r} on {c.anchored_text!r}")
```

`add_comment` accepts a `Run`, a `Paragraph` (wraps every run), or a
`(start_run, end_run)` tuple for ranges. Unlike python-docx's
`Comments.add_comment` (which only writes the part-side body),
`docx_plus` writes the three body-side anchors — so "show in document"
actually jumps to the right place.

### Layout: columns and mid-document section breaks

```python
from docx import Document
from docx_plus.layout import (
    enable_distinct_even_odd_headers,
    insert_section_break,
    set_columns,
)

doc = Document()
doc.add_heading("Intro (single-column)", level=1)
split = doc.add_paragraph("Section break here ↓")

new_section = insert_section_break(split, start_type="continuous")
set_columns(new_section, 2, space=720, separator=True)

doc.add_heading("Body (two-column)", level=1)
for _ in range(10):
    doc.add_paragraph("Lorem ipsum…")

enable_distinct_even_odd_headers(doc)  # doc-level settings.xml flag
doc.save("multicol.docx")
```

### Bookmarks + cross-references

```python
from docx import Document
from docx_plus.bookmarks import add_bookmark, add_cross_reference
from docx_plus.fields import mark_fields_dirty

doc = Document()
heading = doc.add_heading("Introduction", level=1)
add_bookmark(heading, "intro_section")

p = doc.add_paragraph("See ")
add_cross_reference(p, bookmark="intro_section", kind="text")
p.add_run(" on page ")
add_cross_reference(p, bookmark="intro_section", kind="page")

mark_fields_dirty(doc)               # Word recalculates REF / PAGEREF
doc.save("xref.docx")
```

### Footnotes and endnotes

```python
from docx import Document
from docx_plus.notes import add_footnote, add_endnote

doc = Document()
p = doc.add_paragraph("This claim has a footnote")
add_footnote(p, "Sourced from internal benchmarks, 2026-05-19.")
add_endnote(p, "Re-validated against external dataset Q3 2026.")
doc.save("notes.docx")
```

The footnotes part (`word/footnotes.xml`) is created on first use and
round-trips with parsed XML — re-opening the saved document and adding
another footnote inherits the existing ids correctly. Edit existing
notes in place via `edit_footnote(doc, id, text)` /
`edit_endnote(doc, id, text)`; the reference marker stays put.

### Line numbering and page borders

```python
from docx import Document
from docx_plus.layout import Border, set_line_numbering, set_page_borders

doc = Document()
set_line_numbering(doc.sections[0], count_by=5, restart="newPage")

rule = Border(style="single", size=8, color="2F5496", space=24)
set_page_borders(
    doc.sections[0], top=rule, bottom=rule, left=rule, right=rule,
)
doc.save("formal.docx")
```

### Publishing: TOC, captions, Table of Figures

```python
from docx import Document
from docx_plus.fields import mark_fields_dirty
from docx_plus.publishing import add_caption, add_table_of_figures, add_toc

doc = Document()
doc.add_heading("Contents", level=1)
add_toc(doc.add_paragraph(), levels=(1, 2))

doc.add_heading("Architecture", level=1)
doc.add_paragraph("High-level diagram below.")
cap = doc.add_paragraph()
add_caption(cap, "Figure ", caption_type="Figure")
cap.add_run(": System overview.")

doc.add_heading("List of Figures", level=1)
add_table_of_figures(doc.add_paragraph(), caption_type="Figure")

mark_fields_dirty(doc)  # Word populates TOC / SEQ / ToF on open
doc.save("paper.docx")
```

## Command line

`docx-plus` installs a console command (also `python -m docx_plus.cli`)
for inspecting and editing documents from a shell:

```console
$ docx-plus inspect report.docx --provenance     # effective formatting per paragraph
$ docx-plus restyle draft.docx --target Heading1 --target Title -o clean.docx
$ docx-plus controls list form.docx --json       # every content control
$ docx-plus controls set form.docx --tag name --value "Ada Lovelace" -o filled.docx
```

Read commands (`inspect`, `controls list`) take `--json`; so does
`restyle`, which emits its resolved target→style-id mapping as JSON.
Mutating commands (`restyle`, `controls set` / `clear`) require
`-o/--output` (or `--in-place`) so the source is never overwritten by
accident. Full
reference: [`docs/cli.md`](https://thomas-villani.github.io/docx-plus/cli/).

## What's next

v0.2 ships the feature modules listed at the top of this README, plus
the in-place expansion (line numbering, page borders, conditional
table-style formatting, comment / note editing, and the publishing
module). v0.3 added **tracked changes** (read/write revision marks) and
the **`docx-plus` CLI** (`inspect` / `restyle` / `controls`).
[`ROADMAP.md`](ROADMAP.md) tracks what comes after: the backlog
holds `STYLEREF` / sequence-field cross-references, w15 threaded
comments (respond / resolve / reopen), content-control data binding to
Custom XML Parts, bibliography (citations + `BIBLIOGRAPHY` field),
glossary placeholder text, and password-protected forms. Open an issue
if your use case needs any of these.

<details>
<summary>Build history (for contributors)</summary>

- **v0.1.0** — complete: foundation (`core/`), style inspection +
  modification + remapping (`styles/`), content controls (`controls/`),
  fields + document protection (`fields/`, `protection/`), and release
  polish (examples, LibreOffice smoke tests, CI doc build).
- **v0.2.0** — complete: `core/parts`, `comments/`, `layout/`,
  `bookmarks/`, `notes/`, plus the in-place expansion (toggle
  properties, in-place edit verbs, line numbering, page borders,
  conditional table styles, and the `publishing/` module).
- **v0.3.0** — complete: tracked changes (`revisions/`) and the
  `docx-plus` command line (`cli/`).

The per-phase log with dates lives in `IMPLEMENTATION.md` §12.

</details>

## Documentation

Full docs (rendered by [MkDocs](https://www.mkdocs.org) +
[mkdocstrings](https://mkdocstrings.github.io)) are published at
<https://thomas-villani.github.io/docx-plus/>.

- [Architecture](https://thomas-villani.github.io/docx-plus/ARCHITECTURE/)
  — module layout, cascade algorithm, schema-strict insertion, error
  hierarchy, invariants
- [API Index](https://thomas-villani.github.io/docx-plus/API/) —
  hand-curated index of every public symbol with links to the
  auto-generated reference
- **Agent skill** for LLM coding agents:
  [`skills/docx-plus/`](skills/docx-plus/SKILL.md) — point Claude Code (or any
  agent) at it to generate `docx_plus` automation. Overview at
  [docs/SKILLS](https://thomas-villani.github.io/docx-plus/SKILLS/)
- [Test Gaps](https://thomas-villani.github.io/docx-plus/TEST_GAPS/) —
  honest accounting of where the test suite has real holes (snapshot
  at end of Phase 5)
- Per-module API reference lives under
  <https://thomas-villani.github.io/docx-plus/reference/>;
  `uv run mkdocs serve` to browse locally.

## License

MIT. Copyright (c) 2026 Tom Villani, PhD. See [`LICENSE`](LICENSE).
