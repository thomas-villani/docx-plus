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

## Surface at end of Phase 3.5

Phases 4 (`controls/`), 5 (`fields/`, `protection/`), and 6
(`examples/`) are stubs. Their `__init__.py` files exist so imports
resolve, but their `__all__` is empty.

### `docx_plus` (top-level package)

| Symbol | Kind | Notes |
|---|---|---|
| `DocxPlusError` | exception | Root of every typed library error. See [`ARCHITECTURE.md` §7](ARCHITECTURE.md#7-error-hierarchy) |
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
| `NSMAP` | dict | The library's pre-bound namespace map (`w`, `w14`, `r`, `mc`, `a`) |
| `el(tag, **attrs)` | function | Create a namespaced element |
| `sub(parent, tag, **attrs)` | function | Create + append a namespaced child |
| `xpath(node, expr)` | function | XPath against `node` with `NSMAP` pre-bound. Use this — `BaseOxmlElement.xpath()` rejects `namespaces=` kwarg |
| `remove(node)` | function | Detach from parent, no-op if already detached |

`core/parts.py` is a Phase 1 stub (`__all__ = []`). Phase 4 will populate it
with package-part / relationship helpers for custom XML parts.

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

The full SPEC §10 list (`assert_style_not_defined`,
`assert_no_orphan_relationships`, `assert_protected`, `assert_field_dirty`,
`count_controls`) is built out lazily as Phase 4–5 work needs each helper.

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
