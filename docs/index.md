# docx_plus

OOXML-level extensions for [python-docx](https://python-docx.readthedocs.io/).

python-docx is an excellent library that stops at a well-defined
boundary. Past that boundary — the style cascade, content controls,
anchored comments, tracked changes, custom numbering, table borders —
the usual answer is a StackOverflow snippet that reaches into
`element._p` and builds raw `lxml` by hand. Everyone writing serious
document automation ends up with a private, half-tested pile of that
code.

`docx_plus` is that pile, done properly: typed, tested against documents
Word itself authored, and schema-strict about where elements are allowed
to go. It **composes with python-docx** rather than replacing it — you
keep your `Document` object and reach for `docx_plus` only where you
need to.

## Install

```bash
pip install docx-plus
```

```bash
uv add docx-plus
```

Requires Python 3.10+. The only dependencies are `python-docx` and
`lxml`. Current release: **v0.5.0**, published 2026-07-27 on
[PyPI](https://pypi.org/project/docx-plus/).

## Capabilities

| Capability | Detail |
|---|---|
| **[Style cascade](ARCHITECTURE.md#2-the-cascade-resolver)** | Effective formatting for any paragraph / run / cell through the full six-layer cascade, with per-field provenance. Create, modify, and remap styles; materialise any of 107 latent Word built-ins. |
| **[Content controls](ARCHITECTURE.md#6-content-controls)** | Text / dropdown / date / checkbox controls via `FormBuilder`; round-trip read and write of values. |
| **[Comments](ARCHITECTURE.md#76-anchored-comments)** | Anchored comments with the body-side range markers python-docx skips, so "show in document" works. Threading — reply / resolve / reopen (v0.4) — plus durable ids and author presence (v0.5). |
| **[Tracked changes](ARCHITECTURE.md#711-tracked-changes)** | Mark runs as insertions / deletions, read revisions with author and timestamp, accept / reject, toggle track-changes mode (v0.3). |
| **[Fields](ARCHITECTURE.md#7-fields-and-protection)** | `PAGE` / `NUMPAGES` / `DATE` and generic complex fields; mark fields dirty so Word recalculates on next open. |
| **[Tables](ARCHITECTURE.md#714-table-formatting)** | Table / row / cell borders and shading, merge and unmerge, `w:hMerge` normalization, direct-formatting reads (v0.5). Conditional `<w:tblStylePr>` branches resolve the way Word applies them — gated on `w:tblLook`, banding gated on a declared band size. |
| **[Numbering](ARCHITECTURE.md#713-custom-numbering)** | Custom bullet and multi-level numbered list definitions, applied and restarted per paragraph (v0.5). |
| **[Layout](ARCHITECTURE.md#77-layout)** | Multi-column sections, mid-document section breaks, distinct even/odd headers, line numbering, page borders (v0.2). |
| **[Bookmarks](ARCHITECTURE.md#78-bookmarks-and-cross-references)** | Paired body markers plus `REF` / `PAGEREF` cross-references (v0.2). |
| **[Notes](ARCHITECTURE.md#79-footnotes-and-endnotes)** | Footnotes and endnotes over the separate `footnotes.xml` / `endnotes.xml` parts; insert and edit in place (v0.2). |
| **[Publishing](ARCHITECTURE.md#710-publishing)** | Table of Contents, figure / table captions via `SEQ`, Table of Figures (v0.2). |
| **[Protection](ARCHITECTURE.md#7-fields-and-protection)** | Form-fill, read-only, comments-only, or tracked-changes enforcement at the document level. |
| **[Lint](reference/lint.md)** | Audit a document for direct formatting fighting the styles, skipped outline levels, hand-typed lists, and whitespace used as layout — then describe the repair as an ordered, serializable plan (v0.6). Read-only throughout: nothing applies a plan. |
| **[Command line](cli.md)** | `docx-plus inspect / restyle / controls / comments / lint / plan / skill` over the library. |

## Where to start

- New to the library? Read the **[Architecture](ARCHITECTURE.md)**
  overview and skim the API index.
- Want the full reference? **[API Index](API.md)** lists every public
  symbol; **[Reference](reference/core-ns.md)** has per-module pages
  with full signatures and docstrings.
- Driving the library from an LLM coding agent? See the **[Agent
  Skill](SKILLS.md)** — point Claude Code (or any agent) at it to generate
  docx_plus automation without hand-feeding the API.
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

## Project status

**v0.5.0** — beta, and shipping. 1,266 tests, 95% coverage, `mypy
--strict` clean with zero ignores. CI runs Python 3.10–3.13 on Linux
plus a Windows job, and a lower-bound dependency job pinned to
`python-docx==1.0.0` / `lxml==4.9.0`.

The API is stable in practice but pre-1.0: breaking changes are possible
on minor versions and are called out in the
[changelog](https://github.com/thomas-villani/docx-plus/blob/main/CHANGELOG.md).

| Release | Shipped |
|---|---|
| v0.1.0 | Foundation (`core/`), style inspection / modification / remapping, content controls, fields, protection |
| v0.2.0 | Comments, layout, bookmarks, notes, `core/parts`; toggle properties, in-place edits, line numbering, page borders, conditional table styles, `publishing/` |
| v0.3.0 | Tracked changes (`revisions/`) and the `docx-plus` CLI (`inspect`, `restyle`, `controls`) |
| v0.4.0 | Threaded comments over `commentsExtended.xml`, and `docx-plus comments` |
| v0.5.0 | Table formatting (`tables/`), custom numbering (`numbering/`), comment durable ids and author presence, the agent skill in the wheel behind `docx-plus skill` |

[`ROADMAP.md`](https://github.com/thomas-villani/docx-plus/blob/main/ROADMAP.md)
is the live record of what is shipped, backlogged, and deliberately
declined. On the backlog: content-control data binding to Custom XML
Parts, bibliography and `BIBLIOGRAPHY` fields, theme writing, glossary
placeholder text, and password-protected forms.

## Contributing

See
[`CONTRIBUTING.md`](https://github.com/thomas-villani/docx-plus/blob/main/CONTRIBUTING.md)
for the development setup, quality gates, and conventions — including
the expectation that new OOXML output is verified against a file Word
itself authored.
