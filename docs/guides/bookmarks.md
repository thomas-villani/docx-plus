# Bookmarks and cross-references

A bookmark is a paired marker around a span of body text. A cross-reference
is a `REF` (text) or `PAGEREF` (page number) field pointing at a bookmark
by name. Module: `docx_plus.bookmarks`.

Bookmarks themselves are not fields, but cross-references **are** — so
dirty the fields before saving.

## Marking a target and pointing at it

```python
from docx_plus.bookmarks import add_bookmark, add_cross_reference, read_bookmarks
from docx_plus.fields import mark_fields_dirty

heading = doc.add_heading("Introduction", level=1)
add_bookmark(heading, "intro_section")

p = doc.add_paragraph("See ")
add_cross_reference(p, bookmark="intro_section", kind="text")   # REF -> heading text
p.add_run(" on page ")
add_cross_reference(p, bookmark="intro_section", kind="page")   # PAGEREF -> page number

mark_fields_dirty(doc)

for b in read_bookmarks(doc):
    print(b.name, b.anchored_text, b.paragraph_index)
```

`add_bookmark(target, name, *, id_registry=None)` — `target` is a `Run`, a
`Paragraph` (needs ≥1 run), or a `(start_run, end_run)` tuple.

!!! warning "Bookmark names have a strict grammar"
    `name` must match `[A-Za-z_][A-Za-z0-9_]{0,39}` — no spaces, no
    punctuation, 40 characters max. Word's UI silently rejects names
    outside this, but raw OOXML accepts them, which produces a
    cross-reference that never resolves and no error anywhere.
    `add_bookmark` raises eagerly instead.

`add_cross_reference(paragraph, *, bookmark, kind="text", hyperlink=True,
number=None, position=False, suppress_non_delimiters=False,
numeric_format=None, preserve_formatting=False)`

- `kind` is `"text"` (`REF`) or `"page"` (`PAGEREF`). The `\h` switch is
  added by default, so the reference is clickable.
- `number` is `"plain"` / `"relative"` / `"full"` for the target's
  paragraph *number*.
- `position=True` resolves to `"above"` or `"below"`.
- The `REF`-only switches raise if paired with `kind="page"`.

## Reading and deleting

- `read_bookmarks(doc)` returns a `list[BookmarkInfo]` with `bookmark_id`,
  `name`, `anchored_text`, and `paragraph_index`.
- `delete_bookmark(doc, name)` removes every bookmark of that name — by
  name, not id, because that is what cross-references key off. Idempotent.

## Batch inserts and generated names

Share a `BookmarkIdRegistry(doc)` via `id_registry=` when adding several at
once.

`BookmarkNameRegistry(doc)` guards duplicate *names* — a duplicate makes a
`REF` ambiguous — and mints hidden anchors via `next_ref_name()`. That is
Word's own `_Ref` + 9-digit form, which stays out of Word's Bookmark
dialog:

```python
from docx_plus.bookmarks import BookmarkNameRegistry

names = BookmarkNameRegistry(doc)
anchor = names.next_ref_name()      # e.g. "_Ref418320715"
```

This is how you make a [caption
referenceable](publishing.md#see-figure-3-referencing-a-caption) — a `REF`
field cannot point at a `SEQ` field, only at a bookmark.

## Referencing without a bookmark

`STYLEREF` is the exception: it resolves against the nearest paragraph in a
given style and needs no anchor at all. It lives in
[`fields`](fields.md#a-running-header-showing-the-current-chapter) and is
the right tool for running headers.

## See also

- [How bookmarks work](../concepts/bookmarks.md)
- [Fields](fields.md) — `mark_fields_dirty`
- Reference: [`bookmarks.anchor`](../reference/bookmarks-anchor.md),
  [`bookmarks.crossref`](../reference/bookmarks-crossref.md),
  [`bookmarks.read`](../reference/bookmarks-read.md),
  [`bookmarks.registry`](../reference/bookmarks-registry.md)
- Example: `bookmarks_and_xrefs.py`
