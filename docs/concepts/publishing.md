# Publishing

`publishing/` composes the existing [fields plumbing](fields.md) into the
long-document primitives that make Word a viable publishing target.
Three helpers, each emitting a single complex field on top of
`core.build_complex_field`:

- `add_toc(paragraph, *, levels=(1, 3), hyperlink=True, page_numbers=True)`
  emits a `TOC` field. The instruction string is assembled from
  kwargs: `\o "lo-hi"` for outline-level range, `\h` for hyperlinked
  entries, the always-present `\z` and `\u` (Word emits both by
  default), and the optional `\n` to suppress page numbers.
- `add_caption(paragraph, label, *, caption_type="Figure", numbering="ARABIC")`
  emits a label text run (`"Figure "`) followed by a `SEQ` complex
  field. Items sharing the same `caption_type` auto-number together;
  the name is the same vocabulary a Table of Figures uses via its
  `\c` switch.
- `add_table_of_figures(paragraph, *, caption_type="Figure", hyperlink=True)`
  emits `TOC \c "<caption_type>"`, structurally a TOC keyed off the
  matching SEQ captions instead of paragraph outline levels.

For the calls, see the [publishing guide](../guides/publishing.md).

## None of the three marks fields dirty

The publishing module respects the [no-cross-imports
invariant](invariants.md#the-invariants) of importing only from `core/`,
and forwarding to `fields/` would violate it. Users pair their publishing
inserts with one explicit `mark_fields_dirty(doc)` call before save — the
docstrings document the contract.

This is the most common way to get a blank TOC.

## Not covered

Bibliography (sources stored in a Custom XML Part, `<w:sdt>`
citations referencing them, a `BIBLIOGRAPHY` field rendering the
list) sits on the `ROADMAP.md` dependency-gated backlog because it
depends on the CXML data-binding subsystem, which is also unbuilt.
