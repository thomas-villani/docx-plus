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

```python
from docx import Document
from docx_plus.styles import resolve_effective_formatting

doc = Document("report.docx")

# "Why is this heading 13pt and blue?" — a question python-docx can't answer,
# because the value is inherited, not set on the paragraph at all.
resolved = resolve_effective_formatting(doc.paragraphs[0], include_provenance=True)

print(resolved.font_size)                # 13.0
print(resolved.provenance["font_size"])  # FormattingSource(layer='paragraphStyle',
                                         #                  style_id='Heading2', ...)
```

```bash
pip install docx-plus      # or: uv add docx-plus
```

Requires Python 3.10+. The only dependencies are `python-docx` and
`lxml`. Current release: **v0.6.0**, published 2026-07-31 on
[PyPI](https://pypi.org/project/docx-plus/).

## Start here

<div class="grid cards" markdown>

- **[Getting started](getting-started.md)**

    Install, your first script, and the seven conventions that apply
    across every module. Read this first.

- **[Guides](guides/index.md)**

    One task-oriented page per capability — styles, forms, comments,
    tracked changes, tables, publishing, linting, and the rest.

- **[Concepts](concepts/index.md)**

    Why the OOXML behaves the way it does, and why the library is shaped
    around it. The cascade, the parts model, the design commitments.

- **[API reference](API.md)**

    Every public symbol, plus per-module pages with full signatures and
    docstrings generated from source.

</div>

Driving the library from an LLM coding agent? Point it at the
[**agent skill**](SKILLS.md) instead of hand-feeding it the API. Working
from a shell or CI? See the [**CLI**](cli.md).

## Capabilities

| Capability | Detail | Guide |
|---|---|---|
| **Style cascade** | Effective formatting for any paragraph / run / cell through the full eight-layer cascade, with per-field provenance. Create, modify, and remap styles; materialise any of 107 latent Word built-ins. | [Styles](guides/styles.md) |
| **Content controls** | Text / dropdown / date / checkbox controls via `FormBuilder`; round-trip read and write of values. | [Forms](guides/forms.md) |
| **Comments** | Anchored comments with the body-side range markers python-docx skips, so "show in document" works. Threading — reply / resolve / reopen — plus durable ids and author presence. | [Comments](guides/comments.md) |
| **Tracked changes** | Mark runs as insertions / deletions, read revisions with author and timestamp, accept / reject, toggle track-changes mode. | [Tracked changes](guides/revisions.md) |
| **Fields** | `PAGE` / `NUMPAGES` / `DATE` / `STYLEREF` and generic complex fields; mark fields dirty so Word recalculates on next open. | [Fields](guides/fields.md) |
| **Tables** | Table / row / cell borders and shading, merge and unmerge, `w:hMerge` normalization, direct-formatting reads. Conditional `<w:tblStylePr>` branches resolve the way Word applies them. | [Tables](guides/tables.md) |
| **Numbering** | Custom bullet and multi-level numbered list definitions, applied and restarted per paragraph. | [Numbering](guides/numbering.md) |
| **Layout** | Multi-column sections, mid-document section breaks, distinct even/odd headers, line numbering, page borders. | [Layout](guides/layout.md) |
| **Bookmarks** | Paired body markers plus `REF` / `PAGEREF` cross-references. | [Bookmarks](guides/bookmarks.md) |
| **Notes** | Footnotes and endnotes over the separate `footnotes.xml` / `endnotes.xml` parts; insert and edit in place. | [Notes](guides/notes.md) |
| **Publishing** | Table of Contents, figure / table captions via `SEQ`, Table of Figures. | [Publishing](guides/publishing.md) |
| **Protection** | Form-fill, read-only, comments-only, or tracked-changes enforcement at the document level. | [Forms](guides/forms.md#locking-the-document) |
| **Lint** | Audit a document for direct formatting fighting the styles, skipped outline levels, hand-typed lists, and whitespace used as layout — then describe the repair as an ordered, serializable plan. Read-only throughout. | [Linting](guides/linting.md) |
| **Command line** | `docx-plus inspect / restyle / controls / comments / lint / plan / skill` over the library. | [CLI](cli.md) |

## Project status

**v0.6.0** — beta, and shipping. 2,043 tests, 96% coverage, `mypy
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
| v0.6.0 | The linter (`lint/`) — 20 rules, profiles, `plan_fixes`, and `docx-plus lint` / `plan`; the cascade resolver corrected against live Word; the document-wide sweep, `stop_below` baselines, `read_fields` |

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

Auditing the project? See the [Test Gaps](TEST_GAPS.md) snapshot.
