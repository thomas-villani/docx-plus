<div align="center">

# docx_plus

**OOXML-level extensions for [python-docx](https://python-docx.readthedocs.io/).**

[![PyPI](https://img.shields.io/pypi/v/docx-plus.svg?logo=pypi&logoColor=white)](https://pypi.org/project/docx-plus/)
[![Python versions](https://img.shields.io/pypi/pyversions/docx-plus.svg?logo=python&logoColor=white)](https://pypi.org/project/docx-plus/)
[![CI](https://github.com/thomas-villani/docx-plus/actions/workflows/ci.yml/badge.svg)](https://github.com/thomas-villani/docx-plus/actions/workflows/ci.yml)
[![Docs](https://img.shields.io/badge/docs-mkdocs--material-blue)](https://thomas-villani.github.io/docx-plus/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](https://github.com/thomas-villani/docx-plus/blob/main/LICENSE)
[![Typed](https://img.shields.io/badge/typing-strict-blue)](https://mypy-lang.org/)
[![Ruff](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/ruff/main/assets/badge/v2.json)](https://github.com/astral-sh/ruff)

[Documentation](https://thomas-villani.github.io/docx-plus/) ·
[API index](https://thomas-villani.github.io/docx-plus/API/) ·
[Architecture](https://thomas-villani.github.io/docx-plus/ARCHITECTURE/) ·
[Changelog](https://github.com/thomas-villani/docx-plus/blob/main/CHANGELOG.md) ·
[Roadmap](https://github.com/thomas-villani/docx-plus/blob/main/ROADMAP.md)

</div>

---

python-docx is an excellent library that stops at a well-defined boundary.
Past that boundary — the style cascade, content controls, anchored
comments, tracked changes, custom numbering, table borders — the usual
answer is a StackOverflow snippet that reaches into `element._p` and
builds raw `lxml` by hand. Everyone writing serious document automation
ends up with a private, half-tested pile of that code.

`docx_plus` is that pile, done properly: typed, tested against documents
Word itself authored, and schema-strict about where elements are allowed
to go. It **composes with python-docx** rather than replacing it — you
keep your `Document` object and reach for `docx_plus` only where you
need to.

```python
from docx import Document
from docx_plus.styles import resolve_effective_formatting

doc = Document("report.docx")

# "Why is this heading 13pt and blue?" — a question python-docx can't answer,
# because the value is inherited, not set on the paragraph at all.
resolved = resolve_effective_formatting(doc.paragraphs[0], include_provenance=True)

print(resolved.font_size)                # 13.0
print(resolved.provenance["font_size"])  # FormattingSource(layer='paragraphStyle',
                                         #                  style_id='Heading2',
                                         #                  chain_depth=0, ...)
```

## Install

```bash
pip install docx-plus
```

```bash
uv add docx-plus
```

Requires Python 3.10+. The only dependencies are `python-docx` and `lxml`.

## What it does

| | Capability | Module |
|---|---|---|
| **Styles** | Resolve the effective formatting of any paragraph / run / cell through the full eight-layer cascade, with per-field provenance. Create, modify, and remap styles; materialise any of **107** latent Word built-ins. | [`styles/`](https://thomas-villani.github.io/docx-plus/ARCHITECTURE/#2-the-cascade-resolver) |
| **Content controls** | Text, dropdown, date, and checkbox controls via `FormBuilder`; read and write their values; round-trip through save / reopen. | [`controls/`](https://thomas-villani.github.io/docx-plus/ARCHITECTURE/#6-content-controls) |
| **Comments** | Anchored comments with the body-side range markers python-docx omits — so Word's "show in document" actually works. Plus threading (reply / resolve / reopen), durable ids, and author presence. | [`comments/`](https://thomas-villani.github.io/docx-plus/ARCHITECTURE/#76-anchored-comments) |
| **Tracked changes** | Mark runs as insertions or deletions, read every revision with author / timestamp / text, accept or reject them, toggle track-changes mode. | [`revisions/`](https://thomas-villani.github.io/docx-plus/ARCHITECTURE/#711-tracked-changes) |
| **Fields** | `PAGE` / `NUMPAGES` / `DATE` and generic complex fields; mark fields dirty so Word recalculates on open. | [`fields/`](https://thomas-villani.github.io/docx-plus/ARCHITECTURE/#7-fields-and-protection) |
| **Tables** | Table / row / cell borders and shading, cell merging and unmerging, `w:hMerge` normalization, direct-formatting reads. | [`tables/`](https://thomas-villani.github.io/docx-plus/ARCHITECTURE/#714-table-formatting) |
| **Numbering** | Custom bullet and multi-level numbered list definitions, applied and restarted per paragraph. | [`numbering/`](https://thomas-villani.github.io/docx-plus/ARCHITECTURE/#713-custom-numbering) |
| **Layout** | Multi-column sections, mid-document section breaks, distinct even/odd headers, line numbering, page borders. | [`layout/`](https://thomas-villani.github.io/docx-plus/ARCHITECTURE/#77-layout) |
| **Bookmarks** | Paired body markers plus `REF` / `PAGEREF` cross-references. | [`bookmarks/`](https://thomas-villani.github.io/docx-plus/ARCHITECTURE/#78-bookmarks-and-cross-references) |
| **Notes** | Footnotes and endnotes over the separate `footnotes.xml` / `endnotes.xml` parts; insert and edit in place. | [`notes/`](https://thomas-villani.github.io/docx-plus/ARCHITECTURE/#79-footnotes-and-endnotes) |
| **Publishing** | Table of Contents, figure / table captions via `SEQ`, Table of Figures. | [`publishing/`](https://thomas-villani.github.io/docx-plus/ARCHITECTURE/#710-publishing) |
| **Protection** | Form-fill, read-only, comments-only, or tracked-changes enforcement at the document level. | [`protection/`](https://thomas-villani.github.io/docx-plus/ARCHITECTURE/#7-fields-and-protection) |
| **Lint** | Audit a document for direct formatting fighting the styles, skipped outline levels, hand-typed lists, and whitespace used as layout — then describe the repair as an ordered, serializable plan. Read-only. | [`lint/`](https://thomas-villani.github.io/docx-plus/ARCHITECTURE/#715-linting-and-the-fix-plan) |
| **CLI** | `docx-plus inspect / restyle / controls / comments / lint / plan / skill` — the library from a shell. | [`cli/`](https://thomas-villani.github.io/docx-plus/cli/) |

## Quickstart

A few of the most-used surfaces. The
[documentation](https://thomas-villani.github.io/docx-plus/) covers all of
them, and every module has runnable examples under
[`docx_plus/examples/`](https://github.com/thomas-villani/docx-plus/tree/main/docx_plus/examples).

### Styles: define once, apply everywhere

```python
from docx import Document
from docx_plus.styles import apply_style, create_style, ensure_style

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
apply_style(doc.add_paragraph("Hello, world"), "BrandHeading")
doc.save("out.docx")
```

This is the Word-native workflow: define a style, apply it. Change the
style later and every paragraph using it follows — unlike direct
formatting, which you have to remember to update everywhere.

Word's built-ins (`Heading1`–`Heading9`, `Title`, `Quote`, `TOC1`–`TOC9`,
`FootnoteText`, …) are **latent**: defined by Word's defaults but absent
from `styles.xml` until used. `ensure_style` materialises them
idempotently, with defaults extracted from real Word-saved samples
rather than guessed:

```python
ensure_style(doc, "Heading1")   # materialises if absent
ensure_style(doc, "Heading1")   # ...no-op the second time
ensure_style(doc, "TOC2")       # 107 built-ins known
```

For documents authored elsewhere, where a style may be named `"Heading 1"`
with a space, `ensure_style(doc, "Heading1", match_existing=True)` finds
the existing definition via case- and space-insensitive matching — or use
[`remap_styles`](https://thomas-villani.github.io/docx-plus/ARCHITECTURE/#4-style-remapping-phase-35)
for document-wide normalisation.

### Forms: build a fillable document

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

fb.save("form.docx")
```

Read and update an existing form's values:

```python
from docx import Document
from docx_plus.controls import read_controls, set_control_value

doc = Document("form.docx")
set_control_value(doc, "full_name", "Ada Lovelace")
doc.save("filled.docx")

values = read_controls(Document("filled.docx"))
print(values["full_name"].value)   # 'Ada Lovelace'
```

### Comments: anchored to the text they're about

```python
from docx import Document
from docx_plus.comments import add_comment, read_comments, reply_to_comment

doc = Document()
p = doc.add_paragraph()
p.add_run("Project Apollo ")
target = p.add_run("ships next quarter")

c = add_comment(target, "Optimistic — let's see what QA says.", author="Alice")
reply_to_comment(doc, c.comment_id, "Agreed, moving to Q3.", author="Bob")

for comment in read_comments(doc):
    print(f"{comment.author}: {comment.text!r} on {comment.anchored_text!r}")
```

`add_comment` accepts a `Run`, a `Paragraph`, or a `(start_run, end_run)`
tuple for ranges. Unlike python-docx's `Comments.add_comment` — which
writes only the part-side body — `docx_plus` writes the three body-side
anchors, so the comment is attached to a real span of text.

### Tracked changes: propose edits, then accept or reject

```python
from docx import Document
from docx_plus.revisions import accept_revision, mark_insertion, read_revisions

doc = Document()
p = doc.add_paragraph("The report is ")
mark_insertion(p.add_run("nearly "), author="Alice")
p.add_run("complete.")

for rev in read_revisions(doc):
    print(rev.revision_type, rev.author, rev.text)   # 'insertion' 'Alice' 'nearly '

accept_revision(doc, rev.revision_id)   # or reject_revision / accept_all_revisions
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
cap = doc.add_paragraph()
add_caption(cap, "Figure ", caption_type="Figure")
cap.add_run(": System overview.")

doc.add_heading("List of Figures", level=1)
add_table_of_figures(doc.add_paragraph(), caption_type="Figure")

mark_fields_dirty(doc)   # Word populates TOC / SEQ / ToF on open
doc.save("paper.docx")
```

## Command line

`docx-plus` installs a console command (also `python -m docx_plus.cli`)
for inspecting and editing documents from a shell:

```console
$ docx-plus inspect report.docx --provenance        # effective formatting per paragraph
$ docx-plus restyle draft.docx --target Heading1 -o clean.docx
$ docx-plus controls list form.docx --json          # every content control
$ docx-plus controls set form.docx --tag name --value "Ada Lovelace" -o filled.docx
$ docx-plus comments list draft.docx --unresolved   # open comment threads
$ docx-plus lint report.docx                        # formatting defects
$ docx-plus plan report.docx                        # what repairing them would change
$ docx-plus skill install                           # drop the agent skill into .claude/skills/
```

Read commands take `--json`. Mutating commands require `-o/--output` (or
an explicit `--in-place`) so the source is never overwritten by accident.
`lint` and `plan` exit `1` when they found something, so either drops
into a CI step directly. Full reference:
[CLI docs](https://thomas-villani.github.io/docx-plus/cli/).

## For AI coding agents

`docx_plus` ships an **agent skill** inside the package — a structured
guide to the API that Claude Code (or any agent that reads skill files)
can load instead of guessing at signatures. `pip install docx-plus` is
enough to get it:

```console
$ docx-plus skill install      # copies it into ./.claude/skills/
```

See [`docx_plus/skill/SKILL.md`](https://github.com/thomas-villani/docx-plus/blob/main/docx_plus/skill/SKILL.md) and the
[skills overview](https://thomas-villani.github.io/docx-plus/SKILLS/).

## Documentation

Full docs are published at
<https://thomas-villani.github.io/docx-plus/>, built with
[MkDocs](https://www.mkdocs.org) and
[mkdocstrings](https://mkdocstrings.github.io).

- **[API index](https://thomas-villani.github.io/docx-plus/API/)** —
  hand-curated index of every public symbol, linked to the generated
  reference.
- **[Architecture](https://thomas-villani.github.io/docx-plus/ARCHITECTURE/)** —
  module layout, the cascade algorithm, schema-strict insertion, the
  error hierarchy, and the invariants the library maintains. Read this
  if you want to know *why* the OOXML looks the way it does.
- **[CLI reference](https://thomas-villani.github.io/docx-plus/cli/)**.
- **[Test gaps](https://thomas-villani.github.io/docx-plus/TEST_GAPS/)** —
  an honest accounting of where the suite has real holes.

## Project status

**v0.5.0**, released 2026-07-27 — beta, and shipping. 2,043 tests,
96% coverage, `mypy --strict` clean with zero ignores. CI runs Python
3.10–3.13 on Linux plus a Windows job, and a lower-bound dependency job
pinned to `python-docx==1.0.0` / `lxml==4.9.0`.

The API is stable in practice but pre-1.0: breaking changes are possible
on minor versions and will be called out in
[`CHANGELOG.md`](https://github.com/thomas-villani/docx-plus/blob/main/CHANGELOG.md).

[`ROADMAP.md`](https://github.com/thomas-villani/docx-plus/blob/main/ROADMAP.md) is the live record of what is shipped,
backlogged, and deliberately declined. Currently on the backlog:
content-control data binding to Custom XML Parts, bibliography and
`BIBLIOGRAPHY` fields, theme writing, glossary placeholder text, and
password-protected forms. If your use case needs
one of these, [open an issue](https://github.com/thomas-villani/docx-plus/issues/new/choose) —
demand reorders the list.

<details>
<summary>Release history</summary>

- **v0.1.0** — foundation (`core/`), style inspection / modification /
  remapping, content controls, fields, and document protection.
- **v0.2.0** — comments, layout, bookmarks, notes, `core/parts`, plus
  toggle properties, in-place edit verbs, line numbering, page borders,
  conditional table styles, and `publishing/`.
- **v0.3.0** — tracked changes (`revisions/`) and the `docx-plus` CLI.
- **v0.4.0** — threaded comments over `commentsExtended.xml`, and
  `docx-plus comments`.
- **v0.5.0** — table formatting (`tables/`), custom numbering
  (`numbering/`), comment durable ids and author presence
  (`commentsIds.xml` / `people.xml`), and the agent skill shipping in
  the wheel behind `docx-plus skill`.

</details>

## Contributing

Contributions are welcome — see
[`CONTRIBUTING.md`](https://github.com/thomas-villani/docx-plus/blob/main/CONTRIBUTING.md) for the development setup, the
quality gates, and the conventions. In short:

```bash
git clone https://github.com/thomas-villani/docx-plus.git
cd docx-plus
uv sync --extra dev
uv run pre-commit install
uv run pytest
```

Bug reports are most useful with a **minimal `.docx`** attached, or the
offending fragment of `word/document.xml`.

Security issues should be reported privately — see
[`SECURITY.md`](https://github.com/thomas-villani/docx-plus/blob/main/SECURITY.md).

## License

MIT. Copyright (c) 2026 Tom Villani, PhD. See [`LICENSE`](https://github.com/thomas-villani/docx-plus/blob/main/LICENSE).
