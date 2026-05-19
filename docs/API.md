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

Output lands in `site/` (gitignored). Phase 6 will wire this into CI so
the site builds on every push to main.

To browse without serving — read source. Every public symbol has a
Google-style docstring (enforced by ruff's `D` ruleset on `docx_plus/`).

---

## Public surface at v0.1

All six phases are complete. The four runnable example scripts in
`docx_plus/examples/` — `inspect_document.py`, `restyle_existing.py`,
`build_form.py`, `populate_form.py` — are end-to-end demonstrations of the
surface below. Start there if you want to see the library in motion before
reading the index.

### `docx_plus` (top-level package)

| Symbol | Kind | Notes |
|---|---|---|
| `DocxPlusError` | exception | Root of every typed library error. See [`ARCHITECTURE.md` §9](ARCHITECTURE.md#9-error-hierarchy) |
| `__version__` | str | `"0.1.0"` |

### `docx_plus.core`

The foundation primitives. Every capability module imports from here only.

| Symbol | Kind | Notes |
|---|---|---|
| `DocxPlusError` | exception | Re-export of the top-level root |
| `IdRegistry(doc)` | class | Per-document SDT `w:id` allocator. See `core/ids.py` |
| `IdRegistry.next()` | method | Issue a fresh 31-bit positive `w:id` |
| `IdRegistry.reserve(value)` | method | Reserve a specific value or raise `DuplicateIdError` |
| `IdRegistry.issued()` | method | Frozenset snapshot of all issued IDs |
| `DuplicateIdError` | exception | Dual-bases: `DocxPlusError, ValueError` |
| `qn(name)` | function | `"w:tag"` → Clark-notation `{namespace}tag` |
| `NSMAP` | dict | The library's pre-bound namespace map (`w`, `w14`, `r`, `mc`, `a`, `xml`) |
| `XML` | str | XML namespace URI (added Phase 5 to make `qn("xml:space")` work for `w:instrText`) |
| `el(tag, **attrs)` | function | Create a namespaced element |
| `sub(parent, tag, **attrs)` | function | Create + append a namespaced child |
| `xpath(node, expr)` | function | XPath against `node` with `NSMAP` pre-bound. Use this — `BaseOxmlElement.xpath()` rejects `namespaces=` kwarg |
| `remove(node)` | function | Detach from parent, no-op if already detached |

`core/parts.py` remains a Phase 1 stub (`__all__ = []`). v0.2 data
binding (Custom XML Parts) will populate it. v0.1 controls and fields
work inline in the document body so no relationship plumbing is needed.

### `docx_plus.styles` — inspection

The cascade resolver. See [`ARCHITECTURE.md` §2](ARCHITECTURE.md#2-the-cascade-resolver)
for the algorithm walkthrough.

| Symbol | Kind | Notes |
|---|---|---|
| `resolve_effective_formatting(target, *, include_provenance=False)` | function | The headline API — walks six cascade layers, returns `ResolvedFormatting` |
| `ResolvedFormatting` | dataclass (frozen) | 28 formatting fields + `partial` + optional `provenance`. SPEC §4 |
| `FormattingSource` | dataclass (frozen) | `layer`, `style_id`, `chain_depth`, `is_toggle_resolved` |
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

#### Properties accepted by `create_style` / `modify_style`

Field names match `ResolvedFormatting` so output round-trips back through
the modifier without translation. Paragraph-level: `alignment`,
`indent_left`, `indent_right`, `indent_first_line`, `spacing_before`,
`spacing_after`, `line_spacing`, `line_spacing_rule`, `keep_with_next`,
`keep_lines`, `page_break_before`, `outline_level`. Run-level:
`font_name`, `font_size`, `bold`, `italic`, `underline`, `strike`,
`color_rgb`, `highlight`, `caps`, `small_caps`, `vanish`, `vert_align`.

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
