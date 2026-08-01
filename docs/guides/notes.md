# Footnotes and endnotes

`docx_plus.notes` inserts footnotes and endnotes, reads them back, and
edits them in place. The reference marker goes inline into the paragraph;
the note body lives in a separate part (`word/footnotes.xml` /
`word/endnotes.xml`) created on first use and round-tripped on reopen.

These are **not fields** — no `mark_fields_dirty` needed.

## Adding

```python
from docx_plus.notes import add_endnote, add_footnote

p = doc.add_paragraph("This claim has support")
ref = add_footnote(p, "Sourced from internal benchmarks, 2026-05-19.")
add_endnote(p, "Re-validated against external dataset Q3 2026.")

print(ref.note_id)
```

`add_footnote(paragraph, text, *, id_registry=None)` returns a
`FootnoteRef` with `.note_id` and `.body_element`; `add_endnote` returns an
`EndnoteRef` with the same shape.

The note body uses Word's `FootnoteText` / `EndnoteText` paragraph style
and `FootnoteReference` / `EndnoteReference` run style for the leading
glyph, so it looks native.

### Adding several at once

Build a `FootnoteIdRegistry(doc)` or `EndnoteIdRegistry(doc)` and pass it
as `id_registry=` to keep ids unique across the batch.

## Reading

```python
from docx_plus.notes import read_endnotes, read_footnotes

for n in read_footnotes(doc):
    print(n.note_id, n.text, n.paragraph_index)
```

Both return a `list[NoteContent]` with `note_id`, `text`, and
`paragraph_index` (where the body-side reference marker sits).

Word's reserved separator entries — ids `-1` and `0`, and anything typed
`separator` or `continuationSeparator` — are filtered out, so you only ever
see user-authored notes.

## Editing in place

```python
from docx_plus.notes import edit_endnote, edit_footnote

edit_footnote(doc, ref.note_id, "Updated source text.")
```

The body is replaced; the inline marker in the document body is untouched,
so the superscript stays exactly where it was.

Reserved ids (`-1`, `0`) raise `ValueError`; an unknown id raises
`NoteNotFoundError` (also a `KeyError`).

## See also

- [How notes work](../concepts/notes.md) — the separate parts and the
  reserved id ranges
- Reference: [`notes.write`](../reference/notes-write.md),
  [`notes.read`](../reference/notes-read.md),
  [`notes.registry`](../reference/notes-registry.md)
- Example: `footnotes_and_endnotes.py`
