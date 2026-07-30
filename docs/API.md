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

## Public surface at v0.5.0

v0.1's six phases, the initial v0.2 cycle (comments, layout, bookmarks
/ cross-references, footnotes / endnotes), the v0.2 in-place
expansion (toggle props, in-place comment / note edits, line numbering,
page borders, conditional table-style formatting, publishing module),
the v0.3 cycle (tracked changes, the `docx-plus` CLI), the v0.4
cycle (threaded comments with resolve / reopen), and the v0.5 cycle
(table formatting, custom numbering, comment durable ids and author
presence, the packaged agent skill) are all complete.
Thirteen runnable example scripts in
`docx_plus/examples/` demonstrate the surface: `inspect_document.py`,
`restyle_existing.py`, `build_form.py`, `populate_form.py`,
`add_comments.py`, `threaded_comments.py`, `multi_column_layout.py`,
`bookmarks_and_xrefs.py`, `footnotes_and_endnotes.py`,
`publishing_layout.py`, `track_changes.py`, `table_formatting.py`,
`custom_numbering.py`. Start there if you
want to see the library in motion before reading the index.

### `docx_plus` (top-level package)

| Symbol | Kind | Notes |
|---|---|---|
| `DocxPlusError` | exception | Root of every typed library error. See [`ARCHITECTURE.md` §9](ARCHITECTURE.md#9-error-hierarchy) |
| `__version__` | str | `"0.5.0"` |

### `docx_plus.core`

The foundation primitives. Every capability module imports from here only.

| Symbol | Kind | Notes |
|---|---|---|
| `DocxPlusError` | exception | Re-export of the top-level root |
| `IdRegistry(doc)` | class | Per-document SDT `w:id` allocator. See `core/ids.py` |
| `IdRegistry.next()` | method | Issue a fresh 31-bit positive `w:id`, chosen at random |
| `IdRegistry.next_sequential()` | method | v0.5. Issue the *lowest* unused id instead — Word's convention for numbering, where a file full of nine-digit ids would be needlessly unreadable. Gap-filling |
| `IdRegistry.reserve(value)` | method | Reserve a specific value or raise `DuplicateIdError` |
| `IdRegistry.issued()` | method | Frozenset snapshot of all issued IDs |
| `ParaIdRegistry(doc)` | class | v0.4. Per-*package* `w14:paraId` allocator — threaded comments key their parent/child links off it, so it seeds from the body plus the comments / footnotes / endnotes parts. `next_hex()` renders the 8-uppercase-hex-digit form |
| `DuplicateIdError` | exception | Dual-bases: `DocxPlusError, ValueError`. `reserve()` on an already-issued value |
| `IdRangeError` | exception | Dual-bases: `DocxPlusError, ValueError`. A reserved id falls outside the 31-bit positive range |
| `qn(name)` | function | `"w:tag"` → Clark-notation `{namespace}tag` |
| `InvalidNamespaceError` | exception | Dual-bases: `DocxPlusError, ValueError`. `qn()` got a malformed name or unknown prefix |
| `NSMAP` | dict | The library's pre-bound *query* namespace map (`w`, `w14`, `w15`, `w16cid`, `r`, `mc`, `a`, `xml`) |
| `W15` | str | The Word 2012 extension namespace URI — `commentsExtended.xml`, `people.xml`. v0.4 |
| `W16CID` | str | The Word 2016 extension namespace URI — `commentsIds.xml`. v0.5 |
| `XML` | str | XML namespace URI (added Phase 5 to make `qn("xml:space")` work for `w:instrText`) |
| `el(tag, **attrs)` | function | Create a namespaced element |
| `sub(parent, tag, **attrs)` | function | Create + append a namespaced child |
| `xpath(node, expr)` | function | XPath against `node` with `NSMAP` pre-bound. Use this — `BaseOxmlElement.xpath()` rejects `namespaces=` kwarg |
| `remove(node)` | function | Detach from parent, no-op if already detached |
| `body_document_for(proxy, *, operation=...)` | function | Resolve the owning main-body `Document` from a python-docx proxy; raises `ValueError` for header/footer proxies. Shared by `comments` / `notes` |
| `build_complex_field(p_element, instruction, initial_text)` | function | Emit the 5-run complex-field sequence (begin / instrText / separate / result / end). Used by `fields/simple.py` and `bookmarks/crossref.py` |
| `build_bookmark(start_anchor, end_anchor, *, bookmark_id, name)` | function | v0.5. Bracket a range with a `bookmarkStart` / `bookmarkEnd` pair. Lives in core so `publishing` can make a caption referenceable — a `REF` field can only point at a bookmark, never at the caption's `SEQ` field |
| `validate_bookmark_name(name, *, arg_name="name")` | function | v0.5. Check Word's bookmark-name grammar. Shared by the three surfaces that accept one, because a name only Word's UI would reject yields a silently unresolved field |
| `BookmarkIdRegistry(doc)` | class | Bookmark `w:id` allocator. Moved here from `bookmarks` in v0.5 and re-exported there |
| `BookmarkNameRegistry(doc)` | class | v0.5. Bookmark *name* allocator — guards duplicates (which make a `REF` ambiguous) and mints hidden `_Ref` + 9-digit anchors via `next_ref_name()` |
| `DuplicateBookmarkNameError` | exception | Dual-bases: `DocxPlusError, ValueError`. `BookmarkNameRegistry.reserve` on a name already in use |
| `insert_before_first_anchor(parent, new_element, anchor_tags)` | function | Schema-strict insertion helper for `settings.xml` mutations. Used by `fields/update.py` and `layout/settings.py` |
| `ordered_insert(parent, child, order)` | function | v0.5. Idempotent schema-ordered insert given the parent's full child sequence — replaces any same-tag sibling. The stronger form of the above; promoted out of `styles/modify.py` so `numbering` can share it |
| `Border` | dataclass (frozen) | The `CT_Border` shape — `style`, `size`, `color`, `space` — shared by page, table, and cell borders. Defined in `layout` in v0.2, moved to `core` in v0.5; `docx_plus.layout.Border` still works |
| `border_attrs(border)` | function | v0.5. Serialize a `Border` to its four OOXML attributes |
| `get_or_create_part(doc, spec)` | function | Return `(part, root_element)` for a separate OOXML part (creates and wires the relationship if absent). v0.2 |
| `PartSpec` | dataclass (frozen) | Identification data for `get_or_create_part`. Use the pre-baked constants below or build your own |
| `COMMENTS_SPEC` | `PartSpec` | `/word/comments.xml` |
| `COMMENTS_EXTENDED_SPEC` | `PartSpec` | `/word/commentsExtended.xml` — comment threading. v0.4 |
| `COMMENTS_IDS_SPEC` | `PartSpec` | `/word/commentsIds.xml` — durable comment ids. v0.5 |
| `PEOPLE_SPEC` | `PartSpec` | `/word/people.xml` — comment author presence. v0.5 |
| `NUMBERING_SPEC` | `PartSpec` | `/word/numbering.xml` — list definitions. v0.5. Needed because `DocumentPart.numbering_part` fabricates through `NumberingPart.new()`, an unimplemented stub that raises `NotImplementedError` |
| `FOOTNOTES_SPEC` | `PartSpec` | `/word/footnotes.xml` |
| `ENDNOTES_SPEC` | `PartSpec` | `/word/endnotes.xml` |
| `CT_COMMENTS_EXTENDED` / `RT_COMMENTS_EXTENDED` | str | Content- and relationship-type URIs for the extended part. Microsoft extensions, absent from python-docx's `CT` / `RT` enums. v0.4 |
| `CT_COMMENTS_IDS` / `RT_COMMENTS_IDS`, `CT_PEOPLE` / `RT_PEOPLE` | str | Same, for the two v0.5 comment side-parts |

### `docx_plus.styles` — inspection

The cascade resolver. See [`ARCHITECTURE.md` §2](ARCHITECTURE.md#2-the-cascade-resolver)
for the algorithm walkthrough.

| Symbol | Kind | Notes |
|---|---|---|
| `resolve_effective_formatting(target, *, include_provenance=False, table_context=None)` | function | The headline API — walks six cascade layers, returns `ResolvedFormatting`. `table_context` overrides the auto-derived cell position for conditional table-style branches |
| `resolve_paragraph_spacing(paragraph)` | function | The vertical space Word actually leaves around a paragraph, after `<w:contextualSpacing>` and Word's space-after/space-before arithmetic. Returns `ParagraphSpacing` |
| `ParagraphSpacing` | dataclass (frozen) | `space_above` / `space_below` (the applied gaps, in twips), `declared_before` / `declared_after`, `contextual_spacing`, `before_suppressed` / `after_suppressed`. One paragraph's `space_below` equals the next one's `space_above` |
| `ResolvedFormatting` | dataclass (frozen) | 35 formatting fields + `partial` + optional `provenance`. SPEC §4. All twelve ECMA-376 17.7.3 toggles are surfaced (`bold`, `italic`, `cs_bold`, `cs_italic`, `caps`, `small_caps`, `strike`, `vanish`, `emboss`, `imprint`, `outline`, `shadow`). `spacing_before` / `spacing_after` are what the cascade *declares*; `contextual_spacing` carries the flag, and `resolve_paragraph_spacing` applies it |
| `FormattingSource` | dataclass (frozen) | `layer`, `style_id`, `chain_depth`, `is_toggle_resolved`. `layer` is one of `docDefaults`, `tableStyle`, `paragraphStyle`, `styleNumbering`, `numbering`, `directParagraph`, `runStyle`, `directRun` |
| `TableContext` | dataclass (frozen) | Cell position within a table — `is_first_row`, `is_last_row`, `is_first_col`, `is_last_col`, the four `is_band*` fields — plus the table's `w:tblLook` gating (`first_row_enabled`, `last_row_enabled`, `first_col_enabled`, `last_col_enabled`, all defaulting to True). Drives `<w:tblStylePr>` branch selection (`firstRow`, `lastRow`, `band1Horz`, …) |
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
| `load_theme(doc)` | function | Read `word/theme/theme1.xml` (`a:clrScheme` + `a:fontScheme`). Returns `None` on missing/malformed |
| `ThemeColors(scheme, fonts={}, mapping={...})` | dataclass (frozen) | Holds the parsed `a:clrScheme`, `a:fontScheme`, and the document's `<w:clrSchemeMapping>` |
| `ThemeColors.base(theme_name)` | method | Lookup color by Word `ST_ThemeColor` name, resolved through `mapping`; returns `None` for unknowns |
| `ThemeColors.font(token)` | method | Lookup typeface by `ST_Theme` font token (`minorHAnsi`, …); returns `None` for unknowns |
| `resolve_theme_color(theme, name, *, tint=None, shade=None)` | function | Translate aliases + apply tint/shade. Returns hex `RRGGBB` |
| `resolve_theme_font(theme, token)` | function | Resolve a `*Theme` font token to its concrete typeface (e.g. `minorHAnsi` → `Calibri`) |
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
| `add_style_reference(paragraph, *, style, search_from_bottom=False, number=None, position=False, suppress_non_delimiters=False, preserve_formatting=True)` | function | v0.5. `STYLEREF` — the text of the nearest paragraph with a given style, re-resolved per page. The one cross-reference needing no bookmark. `style` is the style *name* (`"Heading 1"`), or an `int` outline level |
| `StyleRefNumber` | type alias | `"plain"` / `"relative"` / `"full"` — how much context a `STYLEREF` number carries (`\n` / `\r` / `\w`) |
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

Anchored, threaded comments — the body-side range markers python-docx
skips, the comment body in `comments.xml`, the thread graph in
`commentsExtended.xml`, and — v0.5 — durable ids in `commentsIds.xml`
plus author presence in `people.xml`. Architecture walkthroughs in
[`ARCHITECTURE.md` §7.6](ARCHITECTURE.md#76-anchored-comments) and
[§7.6.2](ARCHITECTURE.md#762-durable-comment-ids-and-author-presence).

| Symbol | Kind | Notes |
|---|---|---|
| `add_comment(target, text, *, author="", initials=None, id_registry=None, para_id_registry=None, durable_id_registry=None)` | function | Anchor a comment to a `Run`, `Paragraph` (≥1 run required), or `(Run, Run)` tuple. Writes `commentRangeStart` / `commentRangeEnd` / the `CommentReference` marker run, the `<w:comment>` body, a `w14:paraId` stamp plus an unresolved `<w15:commentEx>` thread entry (v0.4), and a `w16cid:durableId` entry (v0.5). Does **not** write `people.xml` |
| `reply_to_comment(doc, parent_id, text, *, author="", initials=None, id_registry=None, para_id_registry=None, durable_id_registry=None)` | function | v0.4. Add a reply beneath an existing comment, mirroring the parent's anchor range. Raises `CommentNotFoundError` if `parent_id` is missing |
| `resolve_comment(doc, comment_id)` | function | v0.4. Mark the whole thread containing `comment_id` resolved (`w15:done="1"`) |
| `reopen_comment(doc, comment_id)` | function | v0.4. The inverse — mark the thread unresolved |
| `edit_comment(doc, comment_id, text)` | function | Replace the body text of an existing comment in place. Preserves `w:author` / `w:date` / `w:initials`, the body-side anchors, and the `w14:paraId` that holds the thread together. Raises `CommentNotFoundError` if id missing |
| `delete_comment(doc, comment_id, *, include_replies=True)` | function | Remove every trace (range markers, reference run, body, thread entry). `include_replies=True` (default) also deletes the subtree, as Word does; `False` promotes orphaned replies to roots. Idempotent — missing id is a no-op |
| `clear_all_comments(doc, *, remove_part=False)` | function | Bulk delete every comment, thread entry, and durable id. `remove_part=True` tears down the comments, commentsExtended, and commentsIds parts. Leaves `people.xml` alone. Idempotent on an empty document |
| `read_comments(doc)` | function | List every comment paired with the document text it anchors. Returns `list[AnchoredComment]` |
| `read_threads(doc)` | function | v0.4. The same comments grouped into threads. Returns `list[CommentThread]` |
| `CommentRef` | dataclass (frozen) | `comment_id`, `body_element` — handle returned by `add_comment` |
| `AnchoredComment` | dataclass (frozen) | `comment_id`, `author`, `initials`, `timestamp`, `text`, `anchored_text`, `paragraph_index`, `parent_id`, `resolved`, `durable_id` |
| `CommentThread` | dataclass (frozen) | v0.4. `root`, `replies`, `resolved` |
| `CommentIdRegistry(doc)` | class | Per-document comment-id allocator. Subclasses the internal `_IdRegistryBase` and seeds from the comments part + any orphaned body anchors |
| `DurableIdRegistry(doc)` | class | v0.5. Per-document `w16cid:durableId` allocator, seeded from `commentsIds.xml` alone. **Hex, not decimal** — use `next_hex()`. Verified against a Word-authored file |
| `set_author_presence(doc, author, *, provider_id="None", user_id=None)` | function | v0.5. Write an author's `people.xml` entry. Idempotent per author; an empty name is a no-op and creates no part. `user_id` defaults to `author`. **Opt-in** — `add_comment` never calls it, since a fabricated directory identity is worse than an absent one |
| `read_author_presence(doc)` | function | v0.5. `list[AuthorPresence]` in document order. `[]` when the part is absent |
| `clear_author_presence(doc, *, remove_part=False)` | function | v0.5. Drop every author entry. Not wired into `delete_comment` — Word keeps stale authors, and pruning needs the author ref-counted across surviving comments |
| `AuthorPresence` | dataclass (frozen) | v0.5. `author` (the only join to `comments.xml`), `provider_id`, `user_id` |
| `LOCAL_PROVIDER` | constant | v0.5. `"None"` — what Word writes for an author with no directory behind them |
| `CommentNotFoundError` | exception | Dual-bases: `DocxPlusError, KeyError`. `edit_comment` / `reply_to_comment` / `resolve_comment` on a missing id |
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
| `add_cross_reference(paragraph, *, bookmark, kind="text", hyperlink=True, number=None, position=False, suppress_non_delimiters=False, numeric_format=None, preserve_formatting=False)` | function | Append a `REF` (`kind="text"`) or `PAGEREF` (`kind="page"`) complex field. `\h` appended by default. v0.5 added the switch surface: `number` → `
`/`
`/`\w` (paragraph number), `position` → `\p` (`"above"`/`"below"`), `numeric_format` → `\#`, `preserve_formatting` → `\* MERGEFORMAT`. `bookmark` is validated. Pair with `mark_fields_dirty` |
| `NumberContext` | type alias | v0.5. `"plain"` / `"relative"` / `"full"` for `number` above |
| `BookmarkRef` | dataclass (frozen) | `bookmark_id`, `name`, `start_element`, `end_element` |
| `BookmarkInfo` | dataclass (frozen) | `bookmark_id`, `name`, `anchored_text`, `paragraph_index` |
| `BookmarkIdRegistry(doc)` | class | Per-document bookmark-id allocator. Re-export of `core.BookmarkIdRegistry` since v0.5 |
| `BookmarkNameRegistry(doc)` | class | v0.5. Per-document bookmark-*name* allocator. `next_ref_name()` mints hidden `_Ref` anchors |
| `DuplicateBookmarkNameError` | exception | Dual-bases: `DocxPlusError, ValueError` |
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

### `docx_plus.numbering`

Custom list definitions — v0.5. python-docx has no `CT_AbstractNum` and
no `CT_Lvl`, so it cannot express what a list *looks like*. OOXML splits
a list into a `<w:abstractNum>` definition and a `<w:num>` instance;
paragraphs reference the **instance**, and that indirection is what makes
restarting possible. See
[`ARCHITECTURE.md` §7.13](ARCHITECTURE.md#713-custom-numbering).

| Symbol | Kind | Notes |
|---|---|---|
| `LevelDefinition` | dataclass (frozen) | One outline level: `fmt`, `text`, `start`, `indent`, `hanging`, `justify`, `suffix`, `restart_after`, `font`. Validated against the ECMA-376 simple types at construction |
| `define_list_definition(doc, *, levels, name=None, style_link=None, num_style_link=None, multi_level_type=None, num_registry=None, abstract_registry=None)` | function | The primitive. Writes one `w:abstractNum` plus one `w:num`; returns the `numId` |
| `define_bullet_list(doc, *, levels=1, indent_step=720, hanging=360, ...)` | function | Preset using Word's round / hollow-`o` / square glyph cycle, each with its symbol font |
| `define_numbered_list(doc, *, levels=1, indent_step=720, hanging=360, ...)` | function | Preset using Word's `1.` / `a.` / `i.` format cycle |
| `apply_list(paragraph, num_id, *, level=0)` | function | Write `w:numPr`. Idempotent; does not validate `num_id` (a dangling reference is legal and renders unnumbered) |
| `remove_list(paragraph, *, suppress_style_numbering=False)` | function | Drop the `w:numPr`. The flag writes the `numId="0"` sentinel instead, the only way to suppress numbering a *style* applies |
| `restart_list(paragraph, num_id, *, level=0, start=1, num_registry=None)` | function | Begin a fresh sequence: adds a second `w:num` over the same `w:abstractNum` with a `w:startOverride`, applies it, returns the new `numId` |
| `read_list_definitions(doc)` | function | Every definition in `numbering.xml`. Returns `[]` when the part is absent; never creates it. Note a fresh `Document()` already has nine from python-docx's template |
| `ListDefinition` | dataclass (frozen) | `num_id`, `abstract_id`, `levels`, `name`, `style_link`, `num_style_link`, `multi_level_type`, `start_overrides` |
| `ListLevel` | dataclass (frozen) | Read-side level. Every field is `None` when its element is absent — which Word reads as its own default, not as zero |
| `NumIdRegistry(doc)` | class | `w:numId` allocator, from 1. Allocates lowest-free via `next_sequential()`, matching Word |
| `AbstractNumIdRegistry(doc)` | class | `w:abstractNumId` allocator, from **0** — the one id namespace where zero is legal |
| `InvalidLevelError` | exception | Dual-bases: `DocxPlusError, ValueError`. A bad `numFmt`, an over-deep `%N` placeholder, more than nine levels |
| `ListDefinitionNotFoundError` | exception | Dual-bases: `DocxPlusError, KeyError`. `restart_list` on an unknown `numId` |
| `MAX_LEVELS` | int | `9` — ECMA-376 caps `w:lvl` per definition |

### `docx_plus.publishing`

Long-document publishing primitives — Table of Contents, captions,
Table of Figures. Each helper emits a complex field; pair with
`docx_plus.fields.mark_fields_dirty` so Word populates the result on
next open. Architecture walkthrough in
[`ARCHITECTURE.md` §7.10](ARCHITECTURE.md#710-publishing).

| Symbol | Kind | Notes |
|---|---|---|
| `add_toc(paragraph, *, levels=(1, 3), hyperlink=True, page_numbers=True)` | function | Append a `TOC` complex field. Instruction string matches Word's default ("Insert → Table of Contents") with `\o`, `\h`, `\z`, `\u`, optional `\n` switches |
| `add_caption(paragraph, label, *, caption_type="Figure", numbering="ARABIC", bookmark_name=None, bookmark_id_registry=None)` | function | Label text run + `SEQ <caption_type> \* <numbering>` complex field. `caption_type` must match the `\c` switch on a downstream Table of Figures. v0.5: `bookmark_name` brackets the label + number in a bookmark, which is the **only** way to make the caption referenceable — a `REF` cannot target a `SEQ` field |
| `add_table_of_figures(paragraph, *, caption_type="Figure", hyperlink=True)` | function | Append a `TOC \c "<caption_type>"` complex field that collects matching captions |

### `docx_plus.tables`

Table **appearance** — the half python-docx omits. It models rows,
columns, cells, widths, and a working `_Cell.merge`, but has no
`CT_Border`, `CT_TblBorders`, `CT_TcBorders`, or `CT_Shd` class and
registers none of those tags. New in v0.5. Architecture walkthrough in
[`ARCHITECTURE.md` §7.14](ARCHITECTURE.md#714-table-formatting).

| Symbol | Kind | Notes |
|---|---|---|
| `set_table_borders(table, *, all_edges=None, top=None, bottom=None, left=None, right=None, inside_h=None, inside_v=None)` | function | Write `<w:tblBorders>`. Full replacement, not a merge; naming no edges removes the element. `all_edges` covers all six; an explicit edge overrides it. **`Border.space` is ignored** — Word writes `w:space="0"` on tables and its UI cannot produce anything else |
| `set_cell_borders(cell, *, all_edges=None, top=None, bottom=None, left=None, right=None, tl2br=None, tr2bl=None)` | function | Write `<w:tcBorders>`. Same semantics. `all_edges` covers the four sides only — the diagonals are a "crossed-out cell" mark, never what a caller means by "all borders" |
| `Shading(fill="auto", pattern="clear", color="auto")` | dataclass | Frozen. `fill` is the background, `pattern` an `ST_Shd` value drawn over it, `color` that pattern's foreground. A solid fill is the default `pattern="clear"` with only `fill` set. Validates all three at construction |
| `set_table_shading(table, shading)` | function | Write `<w:shd>` on `<w:tblPr>`. `None` removes it |
| `set_cell_shading(cell, shading)` | function | The same on `<w:tcPr>` |
| `set_row_shading(row, shading)` | function | **`CT_TrPr` has no `w:shd` child** — there is no row-level shading in the format. Writes through to every cell, as Word does. Iterates `<w:tc>` elements, not `Row.cells`, so a spanning cell is visited once |
| `shading_attrs(shading)` | function | Serialize to the `CT_Shd` attribute mapping |
| `merge_cells(start, end)` | function | Thin wrapper over `_Cell.merge`, translating `InvalidSpanError` into `InvalidMergeError`. Returns the top-left cell of the region, which need not be `start` |
| `unmerge_cell(cell)` | function | The inverse, which python-docx lacks entirely — nothing in it removes a `w:gridSpan` or `w:vMerge`. Works from any cell in the region including a vertical continuation. Content stays in the anchor; widths divide evenly, since the originals were summed away by the merge. Idempotent |
| `normalize_horizontal_merges(table, *, discard_content=False)` | function | Rewrite legacy `<w:hMerge>` spans as `<w:gridSpan>`, which is the only form python-docx's grid model understands. Rendering-preserving (verified against Word 2016). Refuses by default to drop text in a continuation cell — invisible in Word, so keeping it would surface hidden content. Returns the number of regions converted |
| `read_table_formatting(table)` | function | `TableFormatting` — style id, table borders/shading, and a `CellFormatting` per `<w:tc>`. **Direct formatting only**; the table-style cascade is not resolved, so a `Table Grid` table reads back with no borders |
| `TableFormatting` / `CellFormatting` | dataclass | Frozen. `CellFormatting.column` is a *grid offset*, not an index into `Row.cells`. One entry per `<w:tc>`, so a merged cell appears once |
| `InvalidMergeError` | exception | `DocxPlusError` + `ValueError` |

### `docx_plus.revisions`

Tracked changes — read, author, and resolve OOXML revision marks
(`w:ins` / `w:del` / move wrappers / property-change markers).
python-docx cannot read or write tracked changes at all; this module
fills the gap. Scoped in `ROADMAP.md` §1 at the repo root.

| Symbol | Kind | Notes |
|---|---|---|
| `enable_track_changes(doc)` | function | Write `<w:trackChanges/>` into `settings.xml` so Word records every subsequent user edit as a revision. Idempotent (normalises a pre-existing element to "on", collapses duplicates) |
| `disable_track_changes(doc)` | function | Remove every `<w:trackChanges/>`. Idempotent. Existing body revision marks are untouched |
| `mark_insertion(target, *, author="", date=None, id_registry=None)` | function | Wrap existing run(s) in `<w:ins>`. `date=None` stamps current UTC (ms precision). Returns `RevisionRef` |
| `mark_deletion(target, *, author="", date=None, id_registry=None)` | function | Wrap existing run(s) in `<w:del>` and retag each `<w:t>` to `<w:delText>`. Returns `RevisionRef` |
| `read_revisions(doc)` | function | Enumerate every revision in document order, each paired with its metadata and affected text. Returns `list[TrackedChange]` |
| `accept_revision(doc, revision_id)` | function | Accept the revision(s) carrying `revision_id`, keeping the recorded edit. Raises `RevisionNotFoundError` if absent |
| `reject_revision(doc, revision_id)` | function | Reject the revision(s) carrying `revision_id`, restoring the prior state. Raises `RevisionNotFoundError` if absent |
| `accept_all_revisions(doc)` | function | Accept every tracked change. Idempotent; resolves innermost-first |
| `reject_all_revisions(doc)` | function | Reject every tracked change. Idempotent; resolves innermost-first |
| `RevisionRef` | dataclass (frozen) | Write-side handle: `revision_id`, `body_element` (the `<w:ins>` / `<w:del>` element) |
| `TrackedChange` | dataclass (frozen) | Read-side result: `revision_id`, `revision_type`, `author`, `timestamp`, `text`, `paragraph_index` |
| `RevisionIdRegistry(doc)` | class | Per-document revision-id allocator. All revision types share one `w:id` namespace; seeds from every revision-bearing element in the body |
| `RevisionType` | type alias | `Literal["insertion", "deletion", "move_from", "move_to", "format_run", "format_paragraph", "paragraph_mark_insertion", "paragraph_mark_deletion"]` |
| `RevisionTarget` | type alias | `Run | Paragraph | tuple[Run, Run]` — same target shapes as `add_comment`; a range must lie within one paragraph |
| `RevisionNotFoundError` | exception | Dual-bases: `DocxPlusError, KeyError`. `accept_revision` / `reject_revision` on a missing id |

### `docx_plus.cli`

The `docx-plus` command-line interface — a thin shell over the library
(each subcommand wraps one tested function). Full reference, including
every subcommand and flag, lives in [`cli.md`](cli.md).

| Symbol | Kind | Notes |
|---|---|---|
| `main(argv=None)` | function | Console entry point (`docx-plus = "docx_plus.cli:main"`; also `python -m docx_plus.cli`). Returns `0` on success, `1` on a handled library/CLI error, `2` when no command was given |
| `build_parser()` | function | Construct the top-level `argparse.ArgumentParser` with every subcommand registered |
| `docx-plus skill path\|list\|show\|install` | command | v0.5. Locate, read, or install the agent skill packaged at `docx_plus/skill/`. The one command that touches no `.docx`, so it takes `--dest` / `--user` / `--force` rather than `-o/--output` |

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
  element so the inherited value resumes. See
  [`ARCHITECTURE.md` §2](ARCHITECTURE.md#2-the-cascade-resolver).
- **Identifiers.** Style IDs (`w:styleId`) — machine-readable, what
  every function takes. Style names (`w:name`) — human-readable, what
  Word's UI shows. The library accepts IDs everywhere; names are a
  reconciliation concern handled by `find_matching_style` /
  `remap_styles`.
