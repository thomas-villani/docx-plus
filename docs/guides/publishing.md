# Publishing

Tables of contents, auto-numbered figure and table captions, and tables of
figures — the long-document plumbing that makes Word a viable publishing
target. Module: `docx_plus.publishing`.

!!! danger "All three are fields"
    A TOC, a caption's number, and a table of figures all render **blank**
    on disk. Call `mark_fields_dirty(doc)` once, after all inserts and
    before `save`, so Word populates them on open. See
    [fields](fields.md).

## Table of contents

```python
from docx_plus.fields import mark_fields_dirty
from docx_plus.publishing import add_toc

doc.add_heading("Contents", level=1)
add_toc(doc.add_paragraph(), levels=(1, 2))    # collect Heading1..Heading2

# ... add the headings the TOC will collect ...

mark_fields_dirty(doc)
```

`add_toc(paragraph, *, levels=(1, 3), hyperlink=True, page_numbers=True,
additional_styles=None)`

- `levels` — inclusive `(lo, hi)` outline range, each in 1–9, `lo <= hi`.
- `hyperlink=True` — entries are clickable (`\h`).
- `page_numbers=False` — drop page numbers (`\n`), for web-style TOCs.
- `additional_styles` — extra `(style_name, level)` pairs to collect beyond
  the Heading set, e.g. `[("Caption", 4)]`.

The TOC collects paragraphs by **outline level**, so headings need to
actually carry one. `doc.add_heading()` and Word's built-in `Heading1`–`9`
styles do; a custom style needs `outline_level` set — see the [styles
guide](styles.md#which-properties-are-accepted).

## Captions and the table of figures

A caption is a label run (`"Figure "`) followed by a `SEQ` auto-numbering
field. A table of figures collects every caption whose type matches its
`\c` switch — so the `caption_type` on `add_caption` **must match** the one
on `add_table_of_figures`.

```python
from docx_plus.publishing import add_caption, add_table_of_figures

cap = doc.add_paragraph()
add_caption(cap, caption_type="Figure")     # label defaults to "Figure "
cap.add_run(": System overview.")           # your descriptive text

# A "List of Figures" elsewhere:
add_table_of_figures(doc.add_paragraph(), caption_type="Figure")

mark_fields_dirty(doc)      # populates the SEQ numbers and the ToF
```

`add_caption(paragraph, label=None, *, caption_type="Figure",
numbering="ARABIC", bookmark_name=None, bookmark_id_registry=None)`

- `label` defaults to `f"{caption_type} "`; pass `""` to suppress the label
  run entirely.
- `numbering` is a Word picture token — `"ARABIC"`, `"ROMAN"`, `"roman"`,
  `"ALPHABETIC"`, ….
- `caption_type` must be a valid `SEQ` identifier (starts with a letter or
  underscore).

!!! note "The caption paragraph is not auto-styled"
    If you want Word's conventional grey italic caption, apply the style
    yourself. `Caption` is latent, so materialise it first:

    ```python
    from docx_plus.styles import apply_style, ensure_style

    ensure_style(doc, "Caption")
    apply_style(cap, "Caption")
    ```

`add_table_of_figures(paragraph, *, caption_type="Figure", hyperlink=True)`.

## "See Figure 3" — referencing a caption

This is the one that trips everyone up: **a `REF` field cannot point at a
`SEQ` field.** It can only point at a bookmark. So a bare caption is not
referenceable — there is nothing to target.

Pass `bookmark_name` to `add_caption` and it brackets the label plus number
for you:

```python
from docx_plus.bookmarks import BookmarkNameRegistry, add_cross_reference
from docx_plus.publishing import add_caption

names = BookmarkNameRegistry(doc)
anchor = names.next_ref_name()          # e.g. "_Ref418320715" — hidden from Word's UI

cap = doc.add_paragraph()
add_caption(cap, caption_type="Figure", bookmark_name=anchor)
cap.add_run(": Architecture overview")  # added AFTER -> stays outside the bookmark

body = doc.add_paragraph("As shown in ")
add_cross_reference(body, bookmark=anchor)                 # -> "Figure 1"
body.add_run(" on page ")
add_cross_reference(body, bookmark=anchor, kind="page")    # -> "1"
body.add_run(" ")
add_cross_reference(body, bookmark=anchor, position=True)  # -> "above"

mark_fields_dirty(doc)
```

The bookmark spans exactly the extent Word's own "Only label and number"
option uses, so the reference reads `Figure 1` rather than the whole
caption text.

**Order matters:** add the description *after* `add_caption`, or it lands
inside the bookmark and the reference picks it up too.

## End to end

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

mark_fields_dirty(doc)      # populate TOC, SEQ, and ToF on open
doc.save("paper.docx")
```

## See also

- [How publishing works](../concepts/publishing.md)
- [Fields](fields.md) — and why `mark_fields_dirty` is not called for you
- [Bookmarks and cross-references](bookmarks.md)
- Reference: [`publishing.toc`](../reference/publishing-toc.md),
  [`publishing.captions`](../reference/publishing-captions.md),
  [`publishing.figures`](../reference/publishing-figures.md)
- Example: `publishing_layout.py`
