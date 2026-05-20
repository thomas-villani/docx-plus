# docx_plus — API Reference

This file is the hand-curated index of every public symbol. The full
reference (signatures, docstrings, source links) is built by
[MkDocs](https://www.mkdocs.org) with the
[mkdocstrings](https://mkdocstrings.github.io) Python handler. Per-module
pages live under [`reference/`](reference/core-ns.md) and are populated from the
Google-style docstrings on each symbol — there is no separate "regenerate
docs" step beyond running `mkdocs`.

## Serving the docs site locally

```bash
uv run mkdocs serve
# or:
mkdocs serve
```

Opens at <http://127.0.0.1:8000>. Live-reloads on file change. The nav
is configured in `mkdocs.yml` at the repo root.

## Building a static site

```bash
uv run mkdocs build
```

Output lands in `site/` (gitignored). CI wires this on every push to
main (Phase 6).

To browse without serving — read source. Every public symbol has a
Google-style docstring (enforced by ruff's `D` ruleset on `docx_plus/`).

---

## Public surface at v0.2

v0.1's six phases, the initial v0.2 cycle (comments, layout, bookmarks
/ cross-references, footnotes / endnotes), and the v0.2 in-place
expansion (toggle props, in-place comment / note edits, line numbering,
page borders, conditional table-style formatting, publishing module)
are all complete. Nine runnable example scripts in
`docx_plus/examples/` demonstrate the surface: `inspect_document.py`,
`restyle_existing.py`, `build_form.py`, `populate_form.py`,
`add_comments.py`, `multi_column_layout.py`, `bookmarks_and_xrefs.py`,
`footnotes_and_endnotes.py`, `publishing_layout.py`. Start there if you
want to see the library in motion before reading the index.

### `docx_plus` (top-level package)

| Symbol | Kind | Notes |
|---|---|---|
| `DocxPlusError` | exception | Root of every typed library error. See [`ARCHITECTURE.md` §9](ARCHITECTURE.md#9-error-hierarchy) |
| `__version__` | str | `"0.2.0"` |

### `docx_plus.core`

The foundation primitives. Every capability module imports from here only.

| Symbol | Kind | Notes |
|---|---|---|
| `DocxPlusError` | exception | Re-export of the top-level root |
| `IdRegistry(doc)` | class | Per-document SDT `w:id` allocator. See `core/ids.py` |
| `IdRegistry.next()` | method | Issue a fresh 31-bit positive `w:id` |
| `IdRegistry.reserve(value)` | method | Reserve a specific value or raise `DuplicateIdError` |
| `IdRegistry.issued()` | method | Frozenset snapshot of all issued IDs |
| `DuplicateIdError` | exception | Dual-bases: `DocxPlusError, ValueError`. `reserve()` on an already-issued value |
| `IdRangeError` | exception | Dual-bases: `DocxPlusError, ValueError`. A reserved id falls outside the 31-bit positive range |
| `qn(name)` | function | `"w:tag"` → Clark-notation `{namespace}tag` |
| `InvalidNamespaceError` | exception | Dual-bases: `DocxPlusError, ValueError`. `qn()` got a malformed name or unknown prefix |
| `NSMAP` | dict | The library's pre-bound namespace map (`w`, `w14`, `r`, `mc`, `a`, `xml`) |
| `XML` | str | XML namespace URI (added Phase 5 to make `qn("xml:space")` work for `w:instrText`) |
| `el(tag, **attrs)` | function | Create a namespaced element |
| `sub(parent, tag, **attrs)` | function | Create + append a namespaced child |
| `xpath(node, expr)` | function | XPath against `node` with `NSMAP` pre-bound. Use this — `BaseOxmlElement.xpath()` rejects `namespaces=` kwarg |
| `remove(node)` | function | Detach from parent, no-op if already detached |
| `build_complex_field(p_element, instruction, initial_text)` | function | Emit the 5-run complex-field sequence (begin / instrText / separate / result / end). Used by `fields/simple.py` and `bookmarks/crossref.py` |
| `insert_before_first_anchor(parent, new_element, anchor_tags)` | function | Schema-strict insertion helper for `settings.xml` mutations. Used by `fields/update.py` and `layout/settings.py` |
| `get_or_create_part(doc, spec)` | function | Return `(part, root_element)` for a separate OOXML part (creates and wires the relationship if absent). v0.2 |
| `PartSpec` | dataclass (frozen) | Identification data for `get_or_create_part`. Use the pre-baked constants below or build your own |
| `COMMENTS_SPEC` | `PartSpec` | `/word/comments.xml` |
| `FOOTNOTES_SPEC` | `PartSpec` | `/word/footnotes.xml` |
| `ENDNOTES_SPEC` | `PartSpec` | `/word/endnotes.xml` |

### `docx_plus.styles` — inspection

The cascade resolver. See [`ARCHITECTURE.md` §2](ARCHITECTURE.md#2-the-cascade-resolver)
for the algorithm walkthrough.

| Symbol | Kind | Notes |
|---|---|---|
| `resolve_effective_formatting(target, *, include_provenance=False, table_context=None)` | function | The headline API — walks six cascade layers, returns `ResolvedFormatting`. `table_context` overrides the auto-derived cell position for conditional table-style branches |
| `ResolvedFormatting` | dataclass (frozen) | 34 formatting fields + `partial` + optional `provenance`. SPEC §4. All twelve ECMA-376 17.7.3 toggles are surfaced (`bold`, `italic`, `cs_bold`, `cs_italic`, `caps`, `small_caps`, `strike`, `vanish`, `emboss`, `imprint`, `outline`, `shadow`) |
| `FormattingSource` | dataclass (frozen) | `layer`, `style_id`, `chain_depth`, `is_toggle_resolved` |
| `TableContext` | dataclass (frozen) | Cell position within a table — `is_first_row`, `is_last_row`, `is_first_col`, `is_last_col`, `is_band_row`, `is_band_col`. Drives `<w:tblStylePr>` branch selection (`firstRow`, `lastRow`, `band1Horz`, …) per ECMA-376 17.7.6.5 |
| `StyleCascadeError` | exception | `basedOn` cycles or depth > 11 |
| `MissingPartError` | exception | Referenced part absent (reserved — currently no caller raises it) |

### `docx_plus.styles` — modification

Style creation, modification, application, removal, and reconciliation.

| Symbol | Kind | Notes |
|---|---|---|
| `create_style(doc, style_id, *, style_type, name, based_on, next_style, linked_style, ui_priority, q_format, custom, **properties)` | function | Define a new style. Raises `StyleExistsError` if `style_id` is already defined |
| `modify_style(doc, style_id, *, if_missing, **properties)` | function | Mutate one or more properties. Merge semantics for `w:ind`/`w:spacing`/`w:rFonts` |
| `apply_style(target, style_id)` | function | Apply by ID to `Paragraph | Run | _Cell`. Raises `StyleNotFoundError` |
| `delete_style(doc, style_id, *, force=False)` | function | Remove. Raises `StyleInUseError` unless `force=True` (leaves dangling refs) |
| `ensure_style(doc, style_id, *, match_existing=False, **defaults_if_creating)` | function | Idempotent. Materialises latent built-ins from `_BUILTIN_STYLES` if absent |
| `find_matching_style(doc, target_id)` | function | Case/space-insensitive lookup against `w:styleId` and `w:name`. See [`ARCHITECTURE.md` §4](ARCHITECTURE.md#4-style-remapping-phase-35) |
| `remap_styles(doc, *, targets=None, mapping=None, create_missing=False)` | function | Bulk reconciliation via four-step fallback. Rewrites body refs only |
| `list_styles(doc, *, style_type=None, include_latent=False)` | function | Enumerate. `include_latent=True` adds built-ins from `_BUILTIN_STYLES` |
| `StyleProxy` | class | Lightweight live wrapper around a `w:style` element |
| `StyleProxy.modify(**properties)` | method | Delegate to `modify_style` |
| `StyleProxy.delete(*, force=False)` | method | Delegate to `delete_style` |
| `StyleInfo` | dataclass | Returned by `list_styles`: `style_id`, `name`, `style_type`, `based_on`, `is_default`, `is_latent` |
| `StyleExistsError` | exception | `create_style` on duplicate ID |
| `StyleNotFoundError` | exception | Referenced ID not defined |
| `StyleInUseError` | exception | `delete_style` without `force=True` on referenced style |
| `UnknownStylePropertyError` | exception | Dual-bases: `DocxPlusError, TypeError`. Unknown `**properties` kwarg |
| `InvalidColorError` | exception | Dual-bases: `DocxPlusError, ValueError`. A `color_rgb` value that isn't valid `RRGGBB` hex |

#### Properties accepted by `create_style` / `modify_style`

Field names match `ResolvedFormatting` so output round-trips back through
the modifier without translation. Paragraph-level: `alignment`,
`indent_left`, `indent_right`, `indent_first_line`, `spacing_before`,
`spacing_after`, `line_spacing`, `line_spacing_rule`, `keep_with_next`,
`keep_lines`, `page_break_before`, `outline_level`. Run-level:
`font_name`, `font_size`, `bold`, `italic`, `underline`, `strike`,
`color_rgb`, `highlight`, `caps`, `small_caps`, `vanish`, `vert_align`.
(`ResolvedFormatting` additionally exposes the complex-script /
decorative toggles `cs_bold`, `cs_italic`, `emboss`, `imprint`,
`outline`, `shadow` on the read side; `create_style` / `modify_style`
do not yet take them as kwargs — they are read-only.)

### `docx_plus.styles.theme`

Read-only theme color resolution. Theme writing is a v0.2 goal.

| Symbol | Kind | Notes |
|---|---|---|
| `load_theme(doc)` | function | Read `word/theme/theme1.xml`. Returns `None` on missing/malformed |
| `ThemeColors(scheme)` | dataclass (frozen) | Holds the parsed `a:clrScheme` |
| `ThemeColors.base(theme_name)` | method | Lookup by Word `ST_ThemeColor` name; returns `None` for unknowns |
| `resolve_theme_color(theme, name, *, tint=None, shade=None)` | function | Translate aliases + apply tint/shade. Returns hex `RRGGBB` |
| `apply_theme_tint(hex_color, tint_byte)` | function | Lighten toward white |
| `apply_theme_shade(hex_color, shade_byte)` | function | Darken toward black |
| `apply_lum_mod(hex_color, lum_mod)` | function | Multiply lightness by per-mille factor (ECMA-376 17.18.40) |
| `apply_lum_off(hex_color, lum_off)` | function | Add to lightness by per-mille factor |
| `ThemeError` | exception | Structurally invalid input to the transforms |

### `docx_plus.controls` — build side

Build content controls (SDTs) and attach them inline to paragraphs.
Architecture walkthrough in [`ARCHITECTURE.md` §6](ARCHITECTURE.md#6-content-controls).

| Symbol | Kind | Notes |
|---|---|---|
| `FormBuilder(document_or_path=None, *, id_registry=None)` | class | Wrap a `Document`, open one from path, or start fresh. On construction: materialises `PlaceholderText` style, verifies `w14` namespace, seeds `IdRegistry` |
| `FormBuilder.doc` | attribute | The underlying python-docx `Document` — use it for ordinary content (headings, paragraphs, tables) |
| `FormBuilder.add_text_control(paragraph, *, tag, alias=None, placeholder=..., multiline=False)` | method | Single- or multi-line text SDT. Returns the `w:sdt` element |
| `FormBuilder.add_dropdown(paragraph, *, tag, items, alias=None, placeholder=..., editable=False)` | method | Dropdown (or combobox if `editable=True`). `items` is `list[str]` or `list[tuple[display, value]]` |
| `FormBuilder.add_date_picker(paragraph, *, tag, alias=None, placeholder=..., date_format="M/d/yyyy", lcid="en-US")` | method | Date picker SDT |
| `FormBuilder.add_checkbox(paragraph, *, tag, alias=None, checked=False)` | method | Checkbox via `w14:checkbox` |
| `FormBuilder.save(path)` | method | Save the wrapped document. Returns the path as `str` |
| `DropdownItem` | type alias | `str | tuple[str, str]` — display-only or `(display, value)` |
| `MissingNamespaceError` | exception | Document root doesn't declare `w14` — `add_checkbox` would emit unrenderable XML |
| `InvalidDropdownItemError` | exception | Dual-bases: `DocxPlusError, TypeError`. An `items` entry that isn't a `str` or `(display, value)` tuple |

### `docx_plus.controls` — read side

| Symbol | Kind | Notes |
|---|---|---|
| `read_controls(doc, *, by="tag")` | function | Returns `dict[str, ControlValue]` keyed by tag (default) or alias |
| `set_control_value(doc, tag, value)` | function | Update one control by tag. Type-dispatched on the control type |
| `clear_control(doc, tag)` | function | Reset to the placeholder state |
| `ControlValue` | dataclass (frozen) | `tag`, `alias`, `control_type`, `value`, `is_placeholder` |
| `ControlType` | type alias | `Literal["text", "dropdown", "combobox", "date", "checkbox"]` |
| `ControlNotFoundError` | exception | Dual-bases: `DocxPlusError, KeyError`. Tag missing |
| `DuplicateTagError` | exception | Dual-bases: `DocxPlusError, ValueError`. Two SDTs share a tag (repeating-section binding is v0.2) |
| `ValueNotInListError` | exception | Dual-bases: `DocxPlusError, ValueError`. Dropdown value matches neither `w:value` nor `w:displayText`. Combobox is exempt — it accepts freeform |
| `ControlTypeError` | exception | Dual-bases: `DocxPlusError, TypeError`. `set_control_value` value type doesn't match the control type |

### `docx_plus.fields`

Complex field insertion (PAGE / DATE / generic) and the
"recalculate on open" flag. Architecture walkthrough in
[`ARCHITECTURE.md` §7](ARCHITECTURE.md#7-fields-and-protection).

| Symbol | Kind | Notes |
|---|---|---|
| `add_page_number_field(paragraph, *, field="PAGE", format=None)` | function | Append a `PAGE` / `NUMPAGES` / `SECTIONPAGES` field. `format` is a field-switch string like `r"\* ARABIC"`. Returns the begin `<w:r>` |
| `add_date_field(paragraph, *, format="MMMM d, yyyy", auto_update=True)` | function | Append a `DATE` (auto-update) or `CREATEDATE` (frozen) field with a Word date-format string |
| `add_field(paragraph, *, instruction, initial_text="")` | function | Generic complex field. Use for `TOC`, `REF`, `MERGEFIELD`, etc. Spaces are normalised around `instruction` |
| `mark_fields_dirty(doc)` | function | Set `w:updateFields val="true"` in `settings.xml`. Idempotent |
| `PageFieldName` | type alias | `Literal["PAGE", "NUMPAGES", "SECTIONPAGES"]` |

### `docx_plus.protection`

Document-level edit-mode enforcement. Unpassworded — v0.1 by design;
password-protected forms are v0.2 (SPEC §1).

| Symbol | Kind | Notes |
|---|---|---|
| `protect_document(doc, *, mode="forms")` | function | Emit `w:documentProtection` with `w:edit=mode` + `w:enforcement="1"`. Idempotent — second call replaces mode |
| `unprotect_document(doc)` | function | Remove protection. Idempotent |
| `is_protected(doc)` | function | Presence predicate (does not introspect mode) |
| `ProtectionMode` | type alias | `Literal["forms", "readOnly", "comments", "trackedChanges"]` |

### `docx_plus.comments`

Anchored comments — the body-side range markers python-docx skips, plus
the comment body in `comments.xml`. Architecture walkthrough in
[`ARCHITECTURE.md` §7.6](ARCHITECTURE.md#76-anchored-comments).

| Symbol | Kind | Notes |
|---|---|---|
| `add_comment(target, text, *, author="", initials=None, id_registry=None)` | function | Anchor a comment to a `Run`, `Paragraph` (≥1 run required), or `(Run, Run)` tuple. Writes `commentRangeStart` / `commentRangeEnd` / the `CommentReference` marker run, plus the `<w:comment>` body |
| `edit_comment(doc, comment_id, text)` | function | Replace the body text of an existing comment in place. Preserves `w:author` / `w:date` / `w:initials` and the body-side anchors. Raises `CommentNotFoundError` if id missing |
| `delete_comment(doc, comment_id)` | function | Remove all four traces (range markers, reference run, body). Idempotent — missing id is a no-op |
| `clear_all_comments(doc)` | function | Bulk delete every comment by routing each id through `delete_comment`. Idempotent on an empty document |
| `read_comments(doc)` | function | List every comment paired with the document text it anchors. Returns `list[AnchoredComment]` |
| `CommentRef` | dataclass (frozen) | `comment_id`, `body_element` — handle returned by `add_comment` |
| `AnchoredComment` | dataclass (frozen) | `comment_id`, `author`, `initials`, `timestamp`, `text`, `anchored_text`, `paragraph_index` |
| `CommentIdRegistry(doc)` | class | Per-document comment-id allocator. Subclasses the internal `_IdRegistryBase` and seeds from the comments part + any orphaned body anchors |
| `CommentNotFoundError` | exception | Dual-bases: `DocxPlusError, KeyError`. `edit_comment` on a missing id |
| `CommentTarget` | type alias | `Run | Paragraph | tuple[Run, Run]` |

### `docx_plus.layout`

Page-layout extras — columns, mid-document section breaks, doc-level
distinct even/odd headers. Architecture walkthrough in
[`ARCHITECTURE.md` §7.7](ARCHITECTURE.md#77-layout).

| Symbol | Kind | Notes |
|---|---|---|
| `set_columns(section, num, *, space=720, separator=False, widths=None)` | function | Emit `<w:cols>` into the section's `sectPr`. Idempotent (replaces existing). `widths` for unequal columns |
| `insert_section_break(paragraph, *, start_type="nextPage")` | function | Split sections at a chosen paragraph. Clones the trailing `sectPr`, sets `<w:type>`. Returns a `Section` proxy wrapping the new section |
| `enable_distinct_even_odd_headers(doc)` | function | Write `<w:evenAndOddHeaders/>` into `settings.xml`. Idempotent. Distinct from per-section `titlePg` (which python-docx already exposes) |
| `disable_distinct_even_odd_headers(doc)` | function | Remove the element. Idempotent |
| `set_line_numbering(section, *, count_by=1, restart="newPage", start=1, distance=None)` | function | Emit `<w:lnNumType>` for marginal line numbers. Idempotent, schema-strict (lands in its ECMA-376 17.6.17 slot) |
| `set_page_borders(section, *, top=None, bottom=None, left=None, right=None)` | function | Emit `<w:pgBorders>` from one `Border` per side. All-None removes the element. Idempotent, schema-strict |
| `Border` | dataclass (frozen) | One side of a page border: `style`, `size` (eighths of a point), `color` (RGB hex or `"auto"`), `space` (twips from text) |
| `SectionStartType` | type alias | `Literal["nextPage", "continuous", "evenPage", "oddPage", "nextColumn"]` |
| `LineNumberRestart` | type alias | `Literal["newPage", "newSection", "continuous"]` |

### `docx_plus.bookmarks`

Bookmarks and cross-references — paired body markers plus `REF` /
`PAGEREF` complex fields. Architecture walkthrough in
[`ARCHITECTURE.md` §7.8](ARCHITECTURE.md#78-bookmarks-and-cross-references).

| Symbol | Kind | Notes |
|---|---|---|
| `add_bookmark(target, name, *, id_registry=None)` | function | Wrap target with `<w:bookmarkStart>` / `<w:bookmarkEnd>`. Validates `name` against `[A-Za-z_][A-Za-z0-9_]{0,39}` |
| `delete_bookmark(doc, name)` | function | Remove every bookmark with the given name. Idempotent |
| `read_bookmarks(doc)` | function | List every bookmark paired with its anchored text. Returns `list[BookmarkInfo]` |
| `add_cross_reference(paragraph, *, bookmark, kind="text", hyperlink=True)` | function | Append a `REF` (`kind="text"`) or `PAGEREF` (`kind="page"`) complex field. `\h` appended by default. Pair with `mark_fields_dirty` so Word recalculates on open |
| `BookmarkRef` | dataclass (frozen) | `bookmark_id`, `name`, `start_element`, `end_element` |
| `BookmarkInfo` | dataclass (frozen) | `bookmark_id`, `name`, `anchored_text`, `paragraph_index` |
| `BookmarkIdRegistry(doc)` | class | Per-document bookmark-id allocator |
| `BookmarkTarget` | type alias | `Run | Paragraph | tuple[Run, Run]` |
| `CrossReferenceKind` | type alias | `Literal["text", "page"]` |

### `docx_plus.notes`

Footnotes and endnotes — insert-only API for v0.2. Architecture
walkthrough in
[`ARCHITECTURE.md` §7.9](ARCHITECTURE.md#79-footnotes-and-endnotes).

| Symbol | Kind | Notes |
|---|---|---|
| `add_footnote(paragraph, text, *, id_registry=None)` | function | Append the body-side `FootnoteReference` marker run and the `<w:footnote>` body in `footnotes.xml`. Returns `FootnoteRef` |
| `add_endnote(paragraph, text, *, id_registry=None)` | function | Same shape as `add_footnote` but for endnotes |
| `edit_footnote(doc, note_id, text)` | function | Replace the body text of an existing footnote in place. Reserved ids (`-1`, `0`) raise `ValueError`; missing ids raise `NoteNotFoundError` |
| `edit_endnote(doc, note_id, text)` | function | Same shape as `edit_footnote` but for endnotes |
| `read_footnotes(doc)` | function | List user-authored footnotes. Returns `list[NoteContent]`; separator entries (ids -1 / 0) are filtered out |
| `read_endnotes(doc)` | function | Same shape as `read_footnotes` |
| `FootnoteRef` | dataclass (frozen) | `note_id`, `body_element` |
| `EndnoteRef` | dataclass (frozen) | `note_id`, `body_element` |
| `NoteContent` | dataclass (frozen) | `note_id`, `text`, `paragraph_index` |
| `FootnoteIdRegistry(doc)` | class | Per-document footnote-id allocator. Ids -1 / 0 are reserved by Word and refused at reserve time (range check) |
| `EndnoteIdRegistry(doc)` | class | Per-document endnote-id allocator. Same reserved-id treatment |
| `NoteNotFoundError` | exception | Dual-bases: `DocxPlusError, KeyError`. `edit_footnote` / `edit_endnote` on a missing id |

### `docx_plus.publishing`

Long-document publishing primitives — Table of Contents, captions,
Table of Figures. Each helper emits a complex field; pair with
`docx_plus.fields.mark_fields_dirty` so Word populates the result on
next open. Architecture walkthrough in
[`ARCHITECTURE.md` §7.10](ARCHITECTURE.md#710-publishing).

| Symbol | Kind | Notes |
|---|---|---|
| `add_toc(paragraph, *, levels=(1, 3), hyperlink=True, page_numbers=True)` | function | Append a `TOC` complex field. Instruction string matches Word's default ("Insert → Table of Contents") with `\o`, `\h`, `\z`, `\u`, optional `\n` switches |
| `add_caption(paragraph, label, *, caption_type="Figure", numbering="ARABIC")` | function | Label text run + `SEQ <caption_type> \* <numbering>` complex field. `caption_type` must match the `\c` switch on a downstream Table of Figures |
| `add_table_of_figures(paragraph, *, caption_type="Figure", hyperlink=True)` | function | Append a `TOC \c "<caption_type>"` complex field that collects matching captions |

---

## Internal modules (not part of the public API)

These exist in source but are deliberately not re-exported from the
top-level package.

### `docx_plus._testing.ooxml_asserts`

Shared test-suite assertion helpers. Internal — referenced from
`tests/` only.

| Symbol | Notes |
|---|---|
| `assert_ids_unique(doc)` | Every `w:id` on `w:sdt` descendants is unique |
| `assert_style_defined(doc, style_id)` | `w:style[@w:styleId=...]` exists in `word/styles.xml` |
| `count_controls(doc, control_type=None)` | Count SDTs in the body; filter by `"text"`/`"dropdown"`/`"combobox"`/`"date"`/`"checkbox"` |
| `assert_protected(doc, mode=None)` | `w:documentProtection` present with `w:enforcement="1"`; optionally validates `w:edit` |
| `assert_field_dirty(doc)` | `w:updateFields val="true"` present in `settings.xml` |

The SPEC §10 list is now mostly populated; `assert_style_not_defined`
and `assert_no_orphan_relationships` remain unwritten (no caller
needs them yet — see `TEST_GAPS.md` N1).

---

## Conventions

- **Units.** `font_size` in points (float). Spacing, indent, line height
  in twips (int) unless `line_spacing_rule == "auto"`, in which case
  `line_spacing` is a multiplier (e.g. `1.15`). Colors as
  `"RRGGBB"` uppercase hex strings without `#`.
- **Toggle properties** (`bold`, `italic`, `caps`, `small_caps`,
  `strike`, `vanish`). `True` writes the element with no `w:val`.
  `False` writes `w:val="false"`. `None` (in `modify_style`) removes the
  element so XOR with the parent style resumes. See
  [`ARCHITECTURE.md` §2](ARCHITECTURE.md#2-the-cascade-resolver).
- **Identifiers.** Style IDs (`w:styleId`) — machine-readable, what
  every function takes. Style names (`w:name`) — human-readable, what
  Word's UI shows. The library accepts IDs everywhere; names are a
  reconciliation concern handled by `find_matching_style` /
  `remap_styles`.
