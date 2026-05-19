# Changelog

All notable changes to `docx_plus` are recorded here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project
uses [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.2.0] — 2026-05-19

Second cycle. Four new capability modules — anchored comments, layout
extras, bookmarks with cross-references, and footnotes / endnotes —
plus a `core/parts.py` foundation for separate OOXML parts.

### Added

- **Anchored comments** (`docx_plus.comments`) — `add_comment`,
  `read_comments`, `delete_comment`, `CommentRef`, `AnchoredComment`,
  `CommentIdRegistry`. Closes the largest python-docx gap: python-docx
  writes the `<w:comment>` body but skips the three body-side anchors
  (`commentRangeStart` / `commentRangeEnd` / the `CommentReference`
  marker run); `add_comment` writes all four, plus creates the comments
  part on first use. Comment threading (w15) deferred to v0.3.
- **Layout extras** (`docx_plus.layout`) — `set_columns` for `<w:cols>`,
  `insert_section_break` for mid-document section breaks (copies the
  trailing `sectPr`'s properties into the chosen paragraph), and
  `enable_distinct_even_odd_headers` / `disable_…` for the doc-level
  `<w:evenAndOddHeaders/>` flag in `settings.xml`.
- **Bookmarks + cross-references** (`docx_plus.bookmarks`) —
  `add_bookmark`, `read_bookmarks`, `delete_bookmark`, plus
  `add_cross_reference` building `REF` / `PAGEREF` complex fields on
  top of `core.build_complex_field`. `BookmarkIdRegistry` lives in its
  own namespace (separate from SDT and comment ids).
- **Footnotes + endnotes** (`docx_plus.notes`) — `add_footnote`,
  `add_endnote`, `read_footnotes`, `read_endnotes`, paired
  `FootnoteIdRegistry` / `EndnoteIdRegistry`. Reserved ids -1 / 0
  (separator / continuationSeparator) are unissuable; `read_*` filters
  separator entries out of results. Insert-only for v0.2; in-place
  edits of existing notes deferred.
- **`core/parts.py` foundation** — `get_or_create_part(doc, spec)` for
  separate OOXML parts (`comments.xml`, `footnotes.xml`,
  `endnotes.xml`). Registers `XmlPart` subclasses for footnote /
  endnote content types with `PartFactory.part_type_for` so existing
  documents round-trip with parsed XML rather than raw blobs.
- **`core.build_complex_field`** — promoted from `fields/simple.py`'s
  private `_build_complex_field` so cross-references and any future
  field-using module can share it without cross-capability imports.
- **`core.insert_before_first_anchor`** — schema-strict insertion
  helper hoisted from `fields/update.py`. Now used by both
  `fields.mark_fields_dirty` and `layout.enable_distinct_even_odd_headers`.
- **Examples** — `add_comments`, `multi_column_layout`,
  `bookmarks_and_xrefs`, `footnotes_and_endnotes`. Smoke-tested in CI.

### Quality gates

- `mypy --strict` clean on all v0.2 modules.
- `ruff check` clean (Google-convention docstrings).
- Coverage gate at ≥90% holds.
- Examples smoke-tested via `tests/test_examples_smoke.py`.

### Deferred to v0.3+

- w15 threaded comments (parent/child replies).
- Layout: line numbering (`w:lnNumType`), page borders (`w:pgBorders`).
- Comment editing — mutate the body of an existing comment in place.
- Footnote / endnote editing — same.
- Cross-references to headings, numbered items, or captions
  (`STYLEREF`, sequence fields).
- Everything else from SPEC §15 unchanged.

## [0.1.0] — 2026-05-19

First public release. The library composes with `python-docx` rather
than replacing it: callers keep their `Document` object and use
`docx_plus` for the operations `python-docx` cannot reach. See
[`SPEC.md`](SPEC.md) for the API contract and the
[docs site](https://thomas-villani.github.io/docx-plus/) for the full
reference.

### Added

- **Style cascade** (`docx_plus.styles`) — `resolve_effective_formatting`
  walks the six OOXML formatting layers (docDefaults, table style,
  paragraph style chain, numbering, direct paragraph, direct run) and
  returns a fully-resolved `ResolvedFormatting` with optional
  per-field provenance. Cycle detection and depth limit (11) enforced.
- **Style modification** — `create_style`, `modify_style`, `apply_style`,
  `delete_style`, `ensure_style`, `list_styles`. Property kwargs share
  field names with `ResolvedFormatting` for round-trip. Schema-strict
  child ordering enforced on `w:style`, `w:pPr`, `w:rPr`.
- **Style remapping** — `find_matching_style`, `remap_styles` for
  reconciling documents whose style ids differ in casing or spacing
  from the canonical Word ids (`"Heading 1"` vs `"Heading1"`).
- **Latent built-ins** — 107 entries in the built-in styles table
  covering Heading 1–9, Title, Subtitle, Quote, TOC 1–9, body / macro
  / preformatted families, comment and footnote/endnote pairs, table
  defaults, and the common character emphasis set. Defaults extracted
  from real Word-saved samples, not guessed.
- **Theme color resolution** (`docx_plus.styles.theme`) — read-only
  parsing of `theme1.xml` and ECMA-376 17.18.40 tint / shade / lumMod
  / lumOff transforms. Missing or malformed themes set
  `ResolvedFormatting.partial=True` rather than raising.
- **Content controls** (`docx_plus.controls`) — `FormBuilder` writes
  text, dropdown, date picker, and checkbox SDTs inline.
  `read_controls` and `set_control_value` round-trip them through
  save/reopen with type-dispatched value handling.
- **Fields** (`docx_plus.fields`) — `add_page_number_field`,
  `add_date_field`, generic `add_field`, and `mark_fields_dirty`
  (sets `w:updateFields` in `settings.xml` so Word recalculates on
  open).
- **Protection** (`docx_plus.protection`) — `protect_document`,
  `unprotect_document`, `is_protected` for `forms` / `readOnly` /
  `comments` / `trackedChanges` modes. Unpassworded (SPEC §1
  non-goal).
- **Typed error hierarchy** — every library-raised error subclasses
  `DocxPlusError`. Errors with builtin analogues (`ValueError`,
  `TypeError`, `KeyError`) multiple-inherit so existing `except`
  clauses still catch them. See SPEC §16 for the full taxonomy.
- **PEP 561 typing marker** — `docx_plus/py.typed` ships so downstream
  `mypy` users see the type hints.
- **Examples** — `docx_plus.examples.inspect_document`,
  `restyle_existing`, `build_form`, `populate_form`. Runnable as
  `python -m docx_plus.examples.<name>` and smoke-tested in CI.

### Quality gates

- `mypy --strict` clean on `docx_plus/`.
- `ruff check` clean with `D` (pydocstyle, Google convention) on the
  library; relaxed on tests and examples.
- Coverage gate enforced at ≥90% on `core/`, `styles/`, `controls/`.
- Layer-3 LibreOffice headless smoke tests gated by the
  `requires_libreoffice` pytest marker; run on the Ubuntu/Python 3.13
  CI job.

### Known limitations

- The cascade resolver surfaces six of the twelve toggle properties
  in `ResolvedFormatting` (`bold`, `italic`, `caps`, `small_caps`,
  `strike`, `vanish`). The other six (`bCs`, `iCs`, `emboss`,
  `imprint`, `outline`, `shadow`) are spec-recognised but not yet
  exposed — extend `_TOGGLE_RPR` and `ResolvedFormatting` in v0.2.
- `set_control_value` for dates renders `"M/d/yyyy"` identically to
  Word; other formats fall back to ISO 8601 until Word re-renders the
  field on next open. The canonical value in
  `w:date/@w:fullDate` is always correct.
- Conditional table-style formatting (`w:tblStylePr` for firstRow /
  lastRow / etc.) is recognised in the cascade walker but deferred —
  the table style's base `pPr`/`rPr` is applied without conditional
  branches. Tracked in [`docs/TEST_GAPS.md`](docs/TEST_GAPS.md) N2.

### Deferred to v0.2

See SPEC §15 for the full list. Highlights: section / header / footer
first-class API, anchored comments, footnotes / endnotes, bookmarks
and cross-references, table cell shading / borders, theme writing,
password-protected forms, content-control binding to Custom XML Parts.

[Unreleased]: https://github.com/thomas-villani/docx-plus/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/thomas-villani/docx-plus/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/thomas-villani/docx-plus/releases/tag/v0.1.0
