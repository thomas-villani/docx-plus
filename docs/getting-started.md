# Getting started

## Install

```bash
pip install docx-plus
```

```bash
uv add docx-plus
```

Requires Python 3.10+. The only dependencies are `python-docx` and `lxml`.

The import name is `docx_plus` (underscore); the distribution name is
`docx-plus` (hyphen). It depends on `python-docx`, which imports as `docx`.

## Your first script

`docx_plus` **composes with python-docx** — it does not replace it. You
create and own a `docx.Document`, add ordinary content with python-docx as
usual, and call `docx_plus` for the things python-docx can't reach.

```python
from docx import Document
from docx_plus.notes import add_footnote

doc = Document()                          # python-docx, as normal
p = doc.add_paragraph("Revenue grew 14%") # python-docx, as normal
add_footnote(p, "Source: internal Q3 model.")   # docx_plus
doc.save("out.docx")                      # python-docx, as normal
```

Every `docx_plus` function takes a python-docx object — a `Document`,
`Paragraph`, `Run`, `Section`, or `_Cell` — and mutates it in place. There
is no `DocxPlus` document class to construct, and nothing to convert
between.

## The one question python-docx can't answer

The library's headline capability is the style cascade. python-docx will
tell you what a paragraph *declares*; it cannot tell you what a paragraph
*renders as*, because most formatting is inherited rather than set.

```python
from docx import Document
from docx_plus.styles import resolve_effective_formatting

doc = Document("report.docx")
r = resolve_effective_formatting(doc.paragraphs[0], include_provenance=True)

print(r.font_size)                 # 13.0
print(r.bold)                      # True
print(r.provenance["font_size"])   # FormattingSource(layer='paragraphStyle',
                                   #   style_id='Heading2', chain_depth=0, ...)
```

`provenance` is what makes the answer actionable: it names the cascade
layer and the style that actually set each value. See the [styles
guide](guides/styles.md).

## Seven conventions that apply everywhere

These trip up most first attempts. They hold across every module.

### 1. Fields don't show a value until Word recalculates

Anything that inserts a *field* — `add_toc`, `add_caption`,
`add_table_of_figures`, `add_cross_reference`, `add_page_number_field`,
`add_date_field`, `add_field` — writes an **empty placeholder** to disk.
Word fills it in on open, but only if you ask:

```python
from docx_plus.fields import mark_fields_dirty

mark_fields_dirty(doc)     # once, after all inserts, before save
doc.save("report.docx")
```

Forget this and the TOC, page numbers, and cross-references all render
blank. It is the single most common mistake with this library.

### 2. Units are OOXML's, not Word's UI

| Quantity | Unit | Type |
|---|---|---|
| Font size | points | `float` |
| Spacing, indents, column gaps, border offsets, line-number distance | twips (1440 = 1 inch, 20 = 1pt) | `int` |
| Border thickness | eighths of a point | `int` |
| Colours | `"RRGGBB"` uppercase hex, **no** leading `#` | `str` |

### 3. Style IDs are not style names

Functions take the machine-readable `w:styleId` (`"Heading1"`, no space),
*not* the UI name Word shows (`"Heading 1"`). Documents authored elsewhere
routinely disagree about which is which. Reconcile with
`ensure_style(..., match_existing=True)` or `remap_styles` — see the
[styles guide](guides/styles.md#working-with-someone-elses-document).

### 4. Toggle properties have three states

For `bold`, `italic`, `caps`, `small_caps`, `strike`, `vanish` and friends:

- `True` sets it
- `False` forces it **off** (which is different from unset)
- `None`, in `modify_style`, **removes** the element so the value inherits
  from the parent style again

And toggles do not simply override as they cascade — a bold paragraph
style plus a bold character style renders **not bold**. The [cascade
concept page](concepts/cascade.md#toggle-properties) has the full rule and
the worked cases.

### 5. Share an id registry for batch inserts

Comments, bookmarks, notes, and revisions each take an optional
`id_registry=`. When you add several in one session, build one registry
and pass it to every call so the allocated `w:id`s stay unique:

```python
from docx_plus.comments import CommentIdRegistry, add_comment

reg = CommentIdRegistry(doc)
add_comment(run_a, "First.",  author="Alice", id_registry=reg)
add_comment(para_b, "Second.", author="Bob",  id_registry=reg)
```

### 6. Errors are typed, and also catchable as builtins

Every library error subclasses `docx_plus.DocxPlusError`. Most also
subclass the matching builtin, so ordinary `except` clauses still work:

```python
from docx_plus import DocxPlusError
from docx_plus.styles import StyleNotFoundError   # DocxPlusError
from docx_plus.controls import ControlNotFoundError  # DocxPlusError, KeyError
```

The full table is in [invariants and
errors](concepts/invariants.md#error-hierarchy).

### 7. The target is Word

OOXML is rendered by Microsoft Word, and that is what the library's output
is verified against. LibreOffice and Pages mostly work but are not the
contract.

## Where to go next

- **[Guides](guides/index.md)** — one page per capability, task-first.
  Start here if you know what you want to build.
- **[Concepts](concepts/index.md)** — why the OOXML works the way it does,
  and why the library is shaped around it.
- **[CLI](cli.md)** — `docx-plus inspect / restyle / controls / comments /
  lint / plan` for driving the library from a shell or CI.
- **[Agent skill](SKILLS.md)** — point an LLM coding agent at the packaged
  skill instead of hand-feeding it the API.
- **Runnable examples** — fourteen scripts in
  [`docx_plus/examples/`](https://github.com/thomas-villani/docx-plus/tree/main/docx_plus/examples).
  Run one with `python -m docx_plus.examples.track_changes`.
