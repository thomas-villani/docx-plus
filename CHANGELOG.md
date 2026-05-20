# Changelog

All notable changes to `docx_plus` are recorded here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project
uses [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added — publishing hardening (Session D of issues.md review)

- **H13** — `add_toc` gained an optional `additional_styles` keyword:
  a sequence of `(style_name, level)` pairs that get appended to the
  TOC via the ECMA-376 17.16.5.61 `\t` switch. Originally listed in
  the v0.2 expansion plan but not implemented in the initial cycle.
- **M15** — `add_caption`'s `label` is now optional; omitting it
  defaults to `f"{caption_type} "` (the universal case). The library
  example now uses the shorter `add_caption(p, caption_type="Figure")`
  form. Pass `""` to suppress the label run explicitly.

### Fixed — publishing input validation (Session D of issues.md review)

- **H11 / M16** — `add_caption(caption_type=)` and
  `add_table_of_figures(caption_type=)` now validate against the SEQ
  identifier rule (ASCII letter/underscore start, then letters /
  digits / underscores). `add_caption(numbering=)` validates against
  the ECMA-376 17.16.4.1 format-picture token set. Each rejection
  raises `ValueError` with a clear message. Closes a real injection
  vector where a malicious `caption_type` like `'Figure" \o "1-9'`
  could inject additional switches into the `TOC \c` instruction.
- **H12** — `add_toc(levels=)` is now validated as a two-int tuple in
  the 1..9 outline range with `lo <= hi`. Reversed, out-of-range,
  wrong-arity, and non-int inputs now raise `ValueError` with a
  clear message at function entry instead of producing silently
  malformed TOCs.
- **M14** — `add_caption`'s docstring now explicitly notes that the
  caption paragraph is *not* automatically restyled to Word's
  built-in `Caption` paragraph style. Auto-applying the style was
  rejected as too opinionated; callers who want it should write
  `paragraph.style = doc.styles["Caption"]`.

New module `docx_plus/publishing/_validate.py` holds the shared
validation helpers (`validate_seq_identifier`, `validate_numbering_picture`,
`validate_outline_levels`, `validate_additional_styles`).

### Fixed — error taxonomy + cascade interleaving (Session C of issues.md review)

- **C4** — SPEC §9.7 and §16 amended to formally bless the raw
  `ValueError` / `TypeError` carve-out for argument-shape validation at
  the public surface, matching what ARCHITECTURE §9 already
  documented. Typed `DocxPlusError` subclasses remain required for
  domain failures (lookup miss, cascade limit, malformed structure,
  etc.). SPEC §16's table now also lists the v0.2-expansion errors
  `CommentNotFoundError` and `NoteNotFoundError`.
- **H9** — `_apply_table_style_chain` rewritten to walk the basedOn
  chain once and interleave base + matching conditional branches per
  style level (ancestors first), per ECMA-376 17.7.6.5. Previously
  the helper applied base for the whole chain then conditional for
  the whole chain — a child style's base could not override an
  ancestor style's matching `<w:tblStylePr>` branch. Helper
  `_apply_conditional_table_formatting` removed (folded into
  `_apply_table_style_chain`).
- **H10** — `protection/document.py` now imports the shared
  `insert_before_first_anchor` from `core.oxml` instead of carrying a
  byte-identical local copy. Eliminates drift risk.
- **M1** — `add_field` now raises `ValueError` on empty or
  whitespace-only `instruction`. Previously emitted a structurally
  invalid field that Word silently rendered as blank.
- **M2** — `add_page_number_field(format="")` and whitespace-only
  `format` are now treated the same as `format=None` (no double-space
  in the emitted instruction). `format` is stripped on the way in.

### Fixed — schema / part wiring (Session B of issues.md review)

- **C1** — Fresh `footnotes.xml` / `endnotes.xml` parts are now seeded
  with the two reserved separator entries (`w:id="-1" w:type="separator"`
  and `w:id="0" w:type="continuationSeparator"`) Word expects per
  ECMA-376 17.11.16 / 17.11.7. Without them, Word may surface
  "needs repair" prompts and strict consumers may reject the file. The
  `read_footnotes` / `read_endnotes` filter already excludes ids ≤ 0,
  so user-visible note iteration is unchanged.
- **C3** — `<w:pgBorders>` child elements are now written in the
  schema-required sequence `top → left → bottom → right` per
  ECMA-376 17.6.10. Previous order was `top, bottom, left, right` —
  permissive consumers accepted it but strict validators rejected.
- **H6** — `edit_comment` and `edit_footnote` / `edit_endnote` now
  strip ALL block-level children before re-appending the new paragraph,
  not just `<w:p>` children. Comments / notes authored elsewhere can
  legally contain `<w:tbl>`, `<w:sdt>`, `<w:customXml>` per
  ECMA-376 17.13.4.2 + EG_BlockLevelElts; the prior filter left those
  siblings next to the new paragraph.
- **H7** — `set_page_borders` now emits `w:offsetFrom="page"` by
  default, matching Word's UI emission. A new `offset_from` keyword
  (`"page"` | `"text"`) lets callers choose. The `Border.space` docstring
  is corrected: the unit is **points** (range 0-31) per ECMA-376
  17.6.10, not twips as previously stated. New `OffsetFrom` literal
  re-exported from `docx_plus.layout`.
- **H8** — `clear_all_comments` is now single-pass O(N+M): one walk over
  the document body removing every range marker / reference regardless
  of id, then one walk over `comments.xml` removing every entry. Prior
  implementation invoked `delete_comment` per comment, repeating the
  full-body scan N times.

### Fixed — cascade correctness (Session A of issues.md review)

- **C2** — Run-level `w:rStyle` now applies *before* direct run rPr per
  ECMA-376 17.3.2.29. Previously the run's own character-style chain ran
  after the direct rPr, so a style-defined property would override a
  direct one. Provenance for run-level `rStyle` is now reported as a new
  `runStyle` layer (distinct from `linkedCharStyle`, which remains the
  paragraph style's `w:link` companion).
- **H1** — Conditional table-style precedence: `_TBL_STYLE_PR_ORDER`
  now lists rows before columns per ECMA-376 17.7.6.5, so at a cell
  matching both `firstRow` and `firstCol` (with no `nwCell` defined)
  the column branch wins, matching Word.
- **H2** — `<w:dstrike>` (double strikethrough) is now read by the
  resolver and surfaced as `ResolvedFormatting.double_strike`. Handled
  as a non-toggle property (last-writer-wins) per ECMA-376 17.7.3 —
  `dstrike` is not in the toggle property list. Independent of `strike`.
- **H4/H5** — Band2 conditional branches (`band2Horz` / `band2Vert`)
  are now reachable: `TableContext` gained `is_band2_row` and
  `is_band2_col` fields, derived as the complement of band1 at the
  default band-size. The resolver now honors
  `<w:tblStyleRowBandSize>` / `<w:tblStyleColBandSize>` when present on
  the table instance's own `<w:tblPr>` (style-chain lookup remains
  deferred — see TableContext docstring).

## [0.2.0] — 2026-05-19

Second cycle. Four new capability modules — anchored comments, layout
extras, bookmarks with cross-references, and footnotes / endnotes —
plus a `core/parts.py` foundation for separate OOXML parts. The
release was extended in-place to also close every published
"Deferred" bullet and add a publishing module (TOC, captions, Table
of Figures); see `notes-v0_2-expansion-scope.md` at repo root.

### Added — initial cycle

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
  separator entries out of results.
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

### Added — in-place expansion

- **Toggle property completion** — `ResolvedFormatting` now surfaces
  all twelve ECMA-376 17.7.3 toggle properties: the original six
  (`bold`, `italic`, `caps`, `small_caps`, `strike`, `vanish`) plus
  the six complex-script / decorative variants (`cs_bold`,
  `cs_italic`, `emboss`, `imprint`, `outline`, `shadow`). Closes the
  v0.1 "Known limitations" bullet.
- **Comment editing** — `edit_comment(doc, id, text)` and
  `clear_all_comments(doc)`. `CommentNotFoundError` (subclasses
  `DocxPlusError, KeyError`) for missing ids. Body-side anchors and
  `<w:comment>` element attributes (`w:author`, `w:date`,
  `w:initials`) are preserved across edits.
- **Note editing** — `edit_footnote(doc, id, text)` and
  `edit_endnote(doc, id, text)`. `NoteNotFoundError` for missing ids.
  Reserved separator ids (`-1`, `0`) raise `ValueError`.
- **Layout: line numbering** (`docx_plus.layout.set_line_numbering`) —
  emits `<w:lnNumType>` with `count_by` / `restart` / `start` /
  `distance`. Idempotent and schema-strict (sectPr child ordering
  per ECMA-376 17.6.17).
- **Layout: page borders** (`docx_plus.layout.set_page_borders` +
  `Border` dataclass) — emits `<w:pgBorders>` with per-side
  `Border(style, size, color, space)`. All-None removes the element.
- **Conditional table-style formatting** — the cascade resolver
  applies `<w:tblStylePr>` branches (`firstRow`, `lastRow`,
  `firstCol`, `lastCol`, `band1Horz`, `band1Vert`, the four corners,
  `wholeTable`) in ECMA-376 17.7.6.5 precedence order. New
  `TableContext` dataclass; auto-derived from a `_Cell`'s position,
  or pass explicitly to query hypothetical positions.
- **`docx_plus.publishing` module** — `add_toc` (Table of Contents),
  `add_caption` (figure / table captions via `SEQ` complex field),
  `add_table_of_figures` (`TOC \c "Figure"`). Composes existing
  `core.build_complex_field`; users call
  `docx_plus.fields.mark_fields_dirty` before save so Word populates
  results on open.
- **Example** — `publishing_layout` demonstrates TOC + captioned
  figures + ToF. Smoke-tested in CI.

### Quality gates

- `mypy --strict` clean on all v0.2 modules.
- `ruff check` clean (Google-convention docstrings).
- Coverage gate at ≥90% holds (project at ~93%).
- Examples smoke-tested via `tests/test_examples_smoke.py`.

### Deferred to v0.3+

- w15 threaded comments (parent / child replies, resolve / reopen).
- `STYLEREF` / sequence-field cross-references to headings, captions,
  numbered items.
- CLI (`restyle` + `inspect` + `controls` subcommands).
- Custom XML Parts data binding for content controls.
- Bibliography (sources, citations, `BIBLIOGRAPHY` field) — rides on
  CXML data binding.
- Tracked changes read / write API.
- Glossary placeholder text for SDTs.
- Password-protected forms (legacy hash algorithm).
- See SPEC §15 for the remaining held-beyond items.

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
