# `docx_plus.tables.merge`

Merging is one of the few table features python-docx implements
properly: `_Cell.merge` grows a rectangular region and refuses L- and
T-shaped selections. **This module does not re-implement it.** What it
adds is the other three-quarters of the story.

- `merge_cells` — the same operation behind a typed error, so callers
  can `except DocxPlusError` uniformly instead of reaching into
  `docx.exceptions`.
- `unmerge_cell` — the inverse, which python-docx has no notion of.
  Nothing in the package removes a `w:gridSpan` or a `w:vMerge`, so a
  merge is otherwise one-way. Works from any cell in the region,
  including a vertical continuation.
- `normalize_horizontal_merges` — rewrites the *other* horizontal-merge
  encoding.

## The two horizontal-merge encodings

OOXML can express a horizontal merge two ways, and python-docx models
only one of them:

| | Encoding | python-docx |
|---|---|---|
| `w:gridSpan` | One `<w:tc>` widened over several grid columns | Understood |
| `w:hMerge` | One `<w:tc>` per column, followers marked as continuations | Ignored |

Word renders both identically — verified against Word 2016, where a
converted file rasterises byte-for-byte the same as its original. But
on an `hMerge` table `Table.cell` hands back cells that look separate
and are not, and `Row.cells` reports a column count Word never draws.
Word's own COM object model has the same blind spot, counting the
underlying `<w:tc>` elements rather than what it lays out.

`normalize_horizontal_merges` converts the second form into the first,
in place. It refuses by default to drop text held in a continuation
cell: that text is invisible in Word, so keeping it would make hidden
content appear and discarding it silently would lose data. Pass
`discard_content=True` to choose.

Architecture walkthrough:
[`ARCHITECTURE.md` §7.14](../ARCHITECTURE.md#714-table-formatting).

::: docx_plus.tables.merge
    options:
      members:
        - merge_cells
        - unmerge_cell
        - normalize_horizontal_merges
        - InvalidMergeError
