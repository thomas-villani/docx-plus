# Publishing — TOC, captions, notes, bookmarks, cross-references, fields

Long-document plumbing: tables of contents, figure/table captions, table of
figures, footnotes, endnotes, bookmarks, cross-references, and the underlying
field helpers.

Modules: `docx_plus.publishing` (TOC / captions / ToF), `docx_plus.notes`
(footnotes / endnotes), `docx_plus.bookmarks` (bookmarks + cross-refs), and
`docx_plus.fields` (page-number / date / generic fields + the dirty flag).

> **The one rule for this whole file:** TOC, captions, table of figures, and
> cross-references are all **complex fields**. They render blank on disk. Call
> `mark_fields_dirty(doc)` once, after all inserts and before `save`, so Word
> populates them on open. Footnotes/endnotes and bookmarks are *not* fields and
> don't need it.

## Fields (`docx_plus.fields`)

```python
from docx_plus.fields import (
    add_page_number_field, add_date_field, add_field, mark_fields_dirty,
)

# Page X of Y in a footer paragraph:
p = doc.sections[0].footer.paragraphs[0]
p.add_run("Page ")
add_page_number_field(p)                       # PAGE
p.add_run(" of ")
add_page_number_field(p, field="NUMPAGES")     # also: "SECTIONPAGES"

# A date that updates on open (vs CREATEDATE, frozen):
add_date_field(doc.add_paragraph(), format="MMMM d, yyyy", auto_update=True)

# Any other complex field (MERGEFIELD, STYLEREF, ...):
add_field(doc.add_paragraph(), instruction=r'MERGEFIELD FirstName \* MERGEFORMAT')

mark_fields_dirty(doc)   # set updateFields=true so Word recalculates on open
```

- `add_page_number_field(paragraph, *, field="PAGE", format=None)` — `field` is
  `"PAGE"` / `"NUMPAGES"` / `"SECTIONPAGES"`. `format` is a field switch like
  `r"\* ARABIC"`.
- `add_date_field(paragraph, *, format="MMMM d, yyyy", auto_update=True)` —
  `auto_update=False` emits a frozen `CREATEDATE` instead of `DATE`.
- `add_field(paragraph, *, instruction, initial_text="")` — generic; spaces are
  normalized around `instruction`.
- `mark_fields_dirty(doc)` — idempotent; call once before saving.

## Table of contents (`docx_plus.publishing.add_toc`)

```python
from docx_plus.publishing import add_toc
from docx_plus.fields import mark_fields_dirty

doc.add_heading("Contents", level=1)
add_toc(doc.add_paragraph(), levels=(1, 2))    # collect Heading1..Heading2
# ... add the headings the TOC will collect ...
mark_fields_dirty(doc)
```

`add_toc(paragraph, *, levels=(1, 3), hyperlink=True, page_numbers=True, additional_styles=None)`

- `levels` — inclusive `(lo, hi)` outline range, each in 1..9, `lo <= hi`.
- `hyperlink=True` — entries are clickable (`\h`).
- `page_numbers=False` — drop page numbers (`\n`), for web-style TOCs.
- `additional_styles` — extra `(style_name, level)` pairs to collect beyond the
  Heading set, e.g. `[("Caption", 4)]`.

## Captions + table of figures

A caption is a label run (`"Figure "`) followed by a `SEQ` auto-numbering field.
A table of figures collects every caption whose type matches its `\c` switch —
so the `caption_type` on `add_caption` **must match** the one on
`add_table_of_figures`.

```python
from docx_plus.publishing import add_caption, add_table_of_figures

# Beneath a figure:
cap = doc.add_paragraph()
add_caption(cap, caption_type="Figure")        # label defaults to "Figure "
cap.add_run(": System overview.")              # add your descriptive text

# A "List of Figures" elsewhere:
add_table_of_figures(doc.add_paragraph(), caption_type="Figure")

mark_fields_dirty(doc)   # populates SEQ numbers and the ToF
```

- `add_caption(paragraph, label=None, *, caption_type="Figure", numbering="ARABIC")`
  — `label` defaults to `f"{caption_type} "`; pass `""` to suppress the label
  run. `numbering` is a Word picture token (`"ARABIC"`, `"ROMAN"`, `"roman"`,
  `"ALPHABETIC"`, …). `caption_type` must be a valid SEQ identifier (letter/
  underscore start). The caption paragraph is *not* auto-styled — apply Word's
  `Caption` style yourself if you want the conventional grey italic:
  `ensure_style(doc, "Caption"); apply_style(cap, "Caption")` (both from
  `docx_plus.styles`; `Caption` is latent, so `ensure_style` materializes it
  first).
- `add_table_of_figures(paragraph, *, caption_type="Figure", hyperlink=True)`.

## Footnotes & endnotes (`docx_plus.notes`)

Insert-only at the call site plus in-place edits. The reference marker is added
inline to the paragraph; the note body lives in a separate part
(`word/footnotes.xml` / `word/endnotes.xml`), created on first use and
round-tripped on reopen. **Not fields** — no `mark_fields_dirty` needed.

```python
from docx_plus.notes import (
    add_footnote, add_endnote, edit_footnote, read_footnotes, read_endnotes,
)

p = doc.add_paragraph("This claim has support")
ref = add_footnote(p, "Sourced from internal benchmarks, 2026-05-19.")
add_endnote(p, "Re-validated against external dataset Q3 2026.")

# Returned ref carries the note id:
print(ref.note_id)

# Read back (separator entries with reserved ids -1/0 are filtered out):
for n in read_footnotes(doc):
    print(n.note_id, n.text, n.paragraph_index)

# Edit in place; the inline marker stays put:
edit_footnote(doc, ref.note_id, "Updated source text.")
```

- `add_footnote(paragraph, text, *, id_registry=None) -> FootnoteRef` and
  `add_endnote(...) -> EndnoteRef` (`.note_id`, `.body_element`).
- `edit_footnote(doc, note_id, text)` / `edit_endnote(...)` — reserved ids
  (`-1`, `0`) raise `ValueError`; an unknown id raises `NoteNotFoundError`.
- `read_footnotes(doc)` / `read_endnotes(doc)` -> `list[NoteContent]`
  (`note_id`, `text`, `paragraph_index`).
- Adding many in one session: build a `FootnoteIdRegistry(doc)` /
  `EndnoteIdRegistry(doc)` and pass it as `id_registry=` to keep ids unique.

## Bookmarks & cross-references (`docx_plus.bookmarks`)

A bookmark is a paired body marker around a target; a cross-reference is a
`REF` (text) or `PAGEREF` (page) field pointing at a bookmark name. Bookmarks
themselves are not fields, but cross-references **are** — so dirty the fields.

```python
from docx_plus.bookmarks import add_bookmark, add_cross_reference, read_bookmarks
from docx_plus.fields import mark_fields_dirty

heading = doc.add_heading("Introduction", level=1)
add_bookmark(heading, "intro_section")     # name: [A-Za-z_][A-Za-z0-9_]{0,39}

p = doc.add_paragraph("See ")
add_cross_reference(p, bookmark="intro_section", kind="text")   # REF -> heading text
p.add_run(" on page ")
add_cross_reference(p, bookmark="intro_section", kind="page")   # PAGEREF -> page number

mark_fields_dirty(doc)

for b in read_bookmarks(doc):
    print(b.name, b.anchored_text, b.paragraph_index)
```

- `add_bookmark(target, name, *, id_registry=None)` — `target` is a `Run`,
  `Paragraph`, or `(start_run, end_run)` tuple. `name` must match
  `[A-Za-z_][A-Za-z0-9_]{0,39}`.
- `add_cross_reference(paragraph, *, bookmark, kind="text", hyperlink=True)` —
  `kind` is `"text"` or `"page"`; `\h` (clickable) added by default.
- `delete_bookmark(doc, name)` — removes every bookmark of that name
  (idempotent).
- `read_bookmarks(doc) -> list[BookmarkInfo]` (`bookmark_id`, `name`,
  `anchored_text`, `paragraph_index`).
- Batch inserts: share a `BookmarkIdRegistry(doc)` via `id_registry=`.

## End-to-end

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
add_caption(cap, caption_type="Figure")
cap.add_run(": System overview.")

doc.add_heading("List of Figures", level=1)
add_table_of_figures(doc.add_paragraph(), caption_type="Figure")

mark_fields_dirty(doc)     # populate TOC, SEQ, and ToF on open
doc.save("paper.docx")
```

See also: `docx_plus/examples/publishing_layout.py`,
`footnotes_and_endnotes.py`, `bookmarks_and_xrefs.py`.
