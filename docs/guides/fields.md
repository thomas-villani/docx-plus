# Fields

A Word field is a live value the application computes — a page number, a
date, the current chapter heading. `docx_plus.fields` inserts them and sets
the flag that makes Word recalculate.

!!! danger "Fields render blank until Word recalculates them"
    Everything on this page writes an **empty placeholder** to disk. Call
    `mark_fields_dirty(doc)` once, after all your inserts and before
    `save`, so Word populates them on open.

    The same applies to [TOCs and captions](publishing.md) and
    [cross-references](bookmarks.md). [Footnotes](notes.md) and bookmarks
    are *not* fields and don't need it.

## Page numbers

```python
from docx_plus.fields import add_page_number_field, mark_fields_dirty

p = doc.sections[0].footer.paragraphs[0]
p.add_run("Page ")
add_page_number_field(p)                        # PAGE
p.add_run(" of ")
add_page_number_field(p, field="NUMPAGES")      # also: "SECTIONPAGES"

mark_fields_dirty(doc)
```

`add_page_number_field(paragraph, *, field="PAGE", format=None)` — `field`
is `"PAGE"`, `"NUMPAGES"`, or `"SECTIONPAGES"`. `format` is a field switch
such as `r"\* ARABIC"`.

## Dates

```python
from docx_plus.fields import add_date_field

add_date_field(doc.add_paragraph(), format="MMMM d, yyyy", auto_update=True)
```

`add_date_field(paragraph, *, format="MMMM d, yyyy", auto_update=True)`.
With `auto_update=False` it emits a frozen `CREATEDATE` instead of a live
`DATE`.

## A running header showing the current chapter

`STYLEREF` re-resolves per page — it is the only cross-reference that needs
no bookmark.

```python
from docx_plus.fields import add_style_reference

header = doc.sections[0].header.paragraphs[0]
header.add_run("Chapter: ")
add_style_reference(header, style="Heading 1")
```

!!! warning "`style` here is the style *name*, not the style id"
    `"Heading 1"`, with the space — not `"Heading1"`. This is the one place
    in the library that takes a name, because that is what the field
    instruction grammar accepts. Passing an `int` uses an outline level
    (1–9) instead.

`add_style_reference(paragraph, *, style, search_from_bottom=False,
number=None, position=False, suppress_non_delimiters=False,
preserve_formatting=True)`. `search_from_bottom` (`\l`) takes the last
match on the page rather than the first.

## Any other field

```python
from docx_plus.fields import add_field

add_field(doc.add_paragraph(),
          instruction=r'MERGEFIELD FirstName \* MERGEFORMAT')
```

`add_field(paragraph, *, instruction, initial_text="")` — spaces are
normalised around `instruction`. Every helper on this page returns the
begin `<w:r>` element, so you can navigate or relocate the field.

## Reading fields back

```python
from docx_plus.fields import read_fields

for f in read_fields(doc):
    print(f.instruction, f.paragraph_index)
```

## Making Word recalculate

```python
from docx_plus.fields import mark_fields_dirty

mark_fields_dirty(doc)
doc.save("report.docx")
```

This writes `<w:updateFields w:val="true"/>` into `settings.xml`. Word
reads the flag on open, recalculates every field, and resets it — it is a
one-shot mechanism, not persistent state. The call is idempotent, so once
before saving is exactly right.

Readers can also force it manually in Word with `Ctrl+A`, `F9`.

## See also

- [How complex fields work](../concepts/fields.md) — the five-run
  sequence, and why `mark_fields_dirty` exists
- [Publishing](publishing.md) — TOC, captions, table of figures
- [Bookmarks and cross-references](bookmarks.md)
- Reference: [`fields.simple`](../reference/fields-simple.md),
  [`fields.read`](../reference/fields-read.md),
  [`fields.update`](../reference/fields-update.md)
