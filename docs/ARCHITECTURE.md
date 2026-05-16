# docx_plus — Architecture

Present-tense reference for how `docx_plus` is laid out and why. This
document describes what currently exists (end of Phase 3.5). The contract
that constrains it is `SPEC.md`; the meta-guidance on how it was built and
how to extend it is `IMPLEMENTATION.md`. Read this when you need to
understand the library's shape; read those when you need to decide what to
add or how.

Audience: a developer extending or debugging `docx_plus` itself, or a user
who wants more than the README before reading source.

---

## §1 Layout

```
docx_plus/
├── __init__.py              # top-level re-exports (DocxPlusError, __version__)
├── core/                    # foundation primitives — every capability depends on these
│   ├── __init__.py          # DocxPlusError (base of all typed errors)
│   ├── ns.py                # namespace constants + qn()
│   ├── oxml.py              # el(), sub(), xpath(), remove()
│   ├── ids.py               # IdRegistry, DuplicateIdError
│   └── parts.py             # package part / relationship helpers (Phase 1 stub; Phase 4 fills it)
├── styles/                  # complete: inspect, modify, theme
│   ├── __init__.py          # re-exports every public symbol from the submodules
│   ├── inspect.py           # resolve_effective_formatting + ResolvedFormatting + FormattingSource
│   ├── modify.py            # create_style, modify_style, apply_style, delete_style,
│   │                        # ensure_style, find_matching_style, remap_styles, list_styles,
│   │                        # StyleProxy, StyleInfo, _BUILTIN_STYLES table
│   └── theme.py             # ThemeColors, load_theme, resolve_theme_color,
│                            # apply_theme_tint, apply_theme_shade, apply_lum_mod, apply_lum_off
├── controls/                # empty stub — Phase 4 target
├── fields/                  # empty stub — Phase 5 target
├── protection/              # empty stub — Phase 5 target
├── examples/                # empty stub — Phase 6 target
└── _testing/                # internal test helpers (not public API)
    └── ooxml_asserts.py     # assert_ids_unique, assert_style_defined
```

The flat structure is deliberate. Each capability (`styles/`, `controls/`,
…) sits as a sibling of `core/`, never deeper. There is no `_internal/`
hidden layer; `_testing/` is the only underscore-prefixed package, and it
is explicitly excluded from the public surface (`docx_plus/_testing/**`
ignores Google-docstyle in `pyproject.toml`).

---

## §2 The cascade resolver

`styles/inspect.py:resolve_effective_formatting` is the algorithmic core of
the library — the thing python-docx most conspicuously lacks. Given a
`Paragraph`, `Run`, or `_Cell`, it walks six layers of OOXML formatting in
precedence order and returns the values that would actually render.

### Six layers, low-to-high precedence

The cascade is walked at `inspect.py:253-317`
(`_apply_paragraph_cascade`):

1. **`docDefaults`** — `w:docDefaults/{w:rPrDefault, w:pPrDefault}` in
   `styles.xml`. Applied by `_apply_doc_defaults` at `inspect.py:337-353`.
2. **Table style** — only if the target is inside a `w:tbl`. The base
   pPr/rPr from each style in the basedOn chain is applied. Applied by
   `_apply_table_style_chain` at `inspect.py:402-422`. **Conditional
   formatting** (`w:tblStylePr` for firstRow/lastRow/etc.) is recognised in
   SPEC §4 step 2 but deferred — see `TEST_GAPS.md` N2.
3. **Paragraph style chain** — the style named by `w:pStyle` plus every
   `w:basedOn` ancestor. Walked by `_collect_style_chain` at
   `inspect.py:376-399`, then applied root-to-leaf so the most-specific
   style wins. Cycle detection and depth limit (11, per Word) live in
   that one function.
4. **Numbering** — if `w:pPr/w:numPr` is present, the corresponding
   `w:abstractNum/w:lvl` from `numbering.xml` is applied. See
   `_apply_numbering` at `inspect.py:425-466`. If the numbering part is
   missing, `MissingPartError` is **not** raised — the part is treated as
   "not yet materialised" (a common pre-Word state) and skipped silently.
5. **Direct paragraph formatting** — `w:pPr` on the paragraph itself,
   including any `w:rPr` nested under it (paragraph-mark formatting).
6. **Direct run formatting** — `w:rPr` on a target `Run`. Run targets
   also pick up the linked character style (`w:link` on the paragraph
   style) before this layer.

### Toggle properties

Six rPr children are toggles in `_TOGGLE_RPR` at `inspect.py:37-44`:
`b`, `i`, `caps`, `smallCaps`, `strike`, `vanish`. (`bCs`, `iCs`,
`emboss`, `imprint`, `outline`, `shadow` are spec'd toggles but not yet
surfaced on `ResolvedFormatting`; `dstrike` is intentionally **not** a
toggle per ECMA-376.)

Toggle semantics live in `_Accumulator.toggle` at `inspect.py:223-238`:

- An element with `w:val` in `("0", "false")` resets parity to false. Any
  subsequent layer that asserts the toggle flips it back on (this is the
  "explicit override" branch of ECMA-376 17.7.3).
- Any other `w:val` (including absent) XORs against the current parity.
  Even count = false, odd count = true.

This produces the right answer for the test cases listed in
`IMPLEMENTATION.md §5`:

- Style defines bold, no further override → bold
- Style A bold, B basedOn A bold → not bold (XOR)
- Style A bold, B basedOn A `w:b w:val="false"` → not bold (reset)
- Direct bold on a non-bold style → bold
- Direct `w:b w:val="false"` on a bold style → not bold

### Theme color resolution

Implemented in `styles/theme.py`. `load_theme(doc)` at `theme.py` reads
`word/theme/theme1.xml` via the document part's `theme` relationship and
returns a `ThemeColors(scheme=...)`. `resolve_theme_color(theme, name,
*, tint=None, shade=None)` translates Word's `ST_ThemeColor` aliases
(`text1`→`dk1`, `background1`→`lt1`, etc. per ECMA-376 17.18.97), looks
up the base hex, then applies `themeTint` (toward white) or `themeShade`
(toward black). `apply_lum_mod` and `apply_lum_off` implement the
finer-grained luminosity transforms (ECMA-376 17.18.40); they are not
wired into the cascade walker yet but are independently tested.

Theme failures are **graceful**: if the theme part is missing, malformed,
or names an unknown color, `_resolve_color` at `inspect.py:605-620` sets
`acc.partial = True` and returns the unresolved theme name. The
`ResolvedFormatting.partial` flag tells the caller to expect best-effort
values. SPEC §4 ("Theme resolution edge cases") and
`IMPLEMENTATION.md §5` ("Theme resolution can fail gracefully") both
require this — turning the inspector into something that raises on
diverse real-world inputs would be a usability regression.

### Provenance

When `include_provenance=True`, the resolver populates `ResolvedFormatting.provenance`
with a `FormattingSource` per resolved field. The same walk that produces
values produces provenance (`_Accumulator` carries both, gated on
`want_provenance`); `test_provenance_does_not_change_values` in
`tests/test_cascade_provenance.py` is the regression guard that the
values returned with the flag off are bit-identical to those with it on.

`FormattingSource` records:

- `layer` — which of the six cascade layers contributed the value
- `style_id` — for `*Style` layers, the lowest style in the basedOn chain
  that actually set the property (not the leaf style, the *resolving*
  style)
- `chain_depth` — how many basedOn hops away from the target
- `is_toggle_resolved` — True when the value came from the XOR chain
  rather than a single explicit assignment

Provenance is the differentiated feature behind the inspector. It is the
basis for any future "why is this paragraph 14pt italic?" tooling.

---

## §3 Schema-strict insertion

OOXML containers (`CT_Style`, `CT_PPr`, `CT_RPr`, `CT_Settings`, …) have
**required child ordering**. Inserting an element in the wrong position
produces a file Word will silently "repair" on open — which sometimes
works, sometimes doesn't, and is always a latent bug.

`styles/modify.py` enforces order via three canonical sequences:

- `_STYLE_CHILD_ORDER` (`modify.py:67-90`) — the children of a `w:style`
  element
- `_PPR_CHILD_ORDER` (`modify.py:92-129`) — the children of `w:pPr`
- `_RPR_CHILD_ORDER` (`modify.py:131-...`) — the children of `w:rPr`

Every write goes through `_ordered_insert(parent, new_child, order)`,
which finds the canonical position and inserts there, rather than
appending. The `test_*_children_ordered_correctly` family in
`tests/test_styles_modify.py:277-340` verifies the invariant after
`create_style`. (Verification after `modify_style` is on the test-gap
list — see `TEST_GAPS.md` I2.)

All element construction goes through `core/oxml.py`'s `el()` and
`sub()`. No bare `lxml.etree.SubElement` or python-docx `OxmlElement`
calls live in capability modules. This is enforced by the import-
invariant test (see §6).

---

## §4 Style remapping (Phase 3.5)

Real-world documents have a long-running mismatch between style IDs (the
`w:styleId` attribute, what code references) and style names (the
`w:name` attribute, what Word shows in its UI). The same logical style
might appear as `Heading1` in one doc, `Heading 1` (with space) in
another, `heading1` in a third, and `HeadingOne` in a fourth. Code that
calls `apply_style(p, "Heading1")` against the second doc fails — not
because the style is missing, but because the ID doesn't match.

`styles/modify.py:find_matching_style` (line 550) does case- and
space-insensitive lookup against both `w:styleId` and `w:name` of every
defined style. It returns the trivial match when the exact ID is
defined, so it is safe to call unconditionally.

`styles/modify.py:remap_styles` (line 585) is the bulk reconciliation.
For each target ID it walks four steps:

1. **Exact match** — if `target_id` is already defined as a `w:styleId`,
   record the trivial mapping and continue
2. **Explicit mapping** — if `mapping[target_id]` is in the caller's
   dict and points at an existing style, use it
3. **Matcher** — call `find_matching_style(doc, target_id)`. If a
   case/space-insensitive match exists, use the existing definition
4. **Create from built-ins** — only if `create_missing=True` and the
   target is in `_BUILTIN_STYLES`, materialise it via
   `_materialise_builtin`

After resolution, body references (`w:pStyle`, `w:rStyle`, `w:tblStyle`)
are rewritten in-place so subsequent `apply_style` calls work without
translation.

Style-to-style references inside `styles.xml` (`w:basedOn`, `w:next`,
`w:link`, `w:numStyleLink`, `w:styleLink`) are **intentionally not
rewritten**. The remap is a non-destructive body-only rewrite — if the
authoring tool chained `MyHeading` as `basedOn="HeadingOne"`, the chain
is preserved. The cascade resolver will follow it correctly because the
matcher feeds the `apply_style` path, not the inheritance walker.

`ensure_style` accepts a `match_existing=False` flag (added in Phase
3.5). With `True`, it consults `find_matching_style` before falling
back to the built-ins / custom-create path. The returned proxy's
`style_id` may differ from the requested one — callers using
`apply_style` should pass `proxy.style_id` or use `remap_styles` for
document-wide normalisation.

---

## §5 Built-in styles table

`_BUILTIN_STYLES` in `styles/modify.py:1154` enumerates 19 of Word's
built-in styles, covering the SPEC §5 "at minimum" set:

| Style ID | Type | Notes |
|---|---|---|
| `Normal` | paragraph | default; ui_priority=0; qFormat |
| `Heading1` – `Heading9` | paragraph | ui_priority=9; qFormat; sizes 16/13/12/11/11/11/11/10/9 pt (Word-2013 defaults) |
| `Title` | paragraph | ui_priority=10; qFormat; 28pt |
| `Subtitle` | paragraph | ui_priority=11; qFormat; italic 11pt |
| `Quote` | paragraph | ui_priority=29; qFormat; italic |
| `IntenseQuote` | paragraph | ui_priority=30; qFormat; italic + bold |
| `ListParagraph` | paragraph | ui_priority=34; qFormat; indent_left=720 |
| `Caption` | paragraph | ui_priority=35; qFormat; italic 9pt |
| `DefaultParagraphFont` | character | default; ui_priority=1 |
| `Hyperlink` | character | ui_priority=99; color #0563C1; underline single |
| `PlaceholderText` | character | ui_priority=99; color #808080 |
| `TableNormal` | table | default; ui_priority=99 |
| `NoList` | numbering | default; ui_priority=99 |

The numbers come from extracting `styles.xml` from documents Word has
itself materialised, not from guessing. Built-ins materialise *without*
`w:customStyle="1"` (they are not user-defined) and the four `default`
entries carry `w:default="1"`.

`ensure_style` is idempotent and aware that **python-docx already ships
a `styles.xml` with most of these latent built-ins materialised**, but
at Word-2007 defaults (e.g. Heading1 = 14pt #365F91), not Word-2013.
This is deliberate: `ensure_style` consults the built-ins table **only**
when the ID is genuinely missing from `styles.xml`. If python-docx
already shipped it, the existing definition is returned unchanged. The
table is a "the style is absent, here is what Word would have written"
fallback, not a "force my preferred defaults" mechanism.

---

## §6 Invariants

These are the architectural commitments. Each is enforced by a test.

1. **No imports between capability modules.** `styles/`, `controls/`,
   `fields/`, `protection/` may import from `core/` only — never from each
   other. Enforced by `tests/test_import_invariant.py`, which walks the
   AST of every `.py` file in each capability directory and asserts no
   import names another capability.

2. **All XML element construction goes through `core/oxml.py`.** No bare
   `lxml.etree.SubElement` or `OxmlElement` calls in capability modules.
   No string-formatted XML anywhere. The convention makes it possible to
   add validation/logging hooks later without rewriting every call site.

3. **`IdRegistry` is the only source of new `w:id` values on SDT
   elements.** Capability modules either receive a registry as a
   parameter or create one scoped to their lifetime. Other ID
   namespaces (`r:id`, `w:bookmarkStart/@w:id`,
   `w:commentRangeStart/@w:id`) are distinct uniqueness domains and will
   get their own registries in later phases.

4. **No magic attributes on python-docx objects.** Library state lives
   in `docx_plus`-owned objects (`IdRegistry`, `StyleProxy`, and in
   Phase 4, `FormBuilder`). Never `setattr(doc, "_my_state", ...)`.

5. **All public functions have type hints.** `mypy --strict` passes on
   `docx_plus/`. The test suite uses looser hints.

6. **All public functions have Google-style docstrings.** Module
   docstring, function summary, Args/Returns/Raises sections. Enforced
   by ruff's `D` ruleset (`pyproject.toml:70-83`); `_testing/`,
   `examples/`, and `tests/` are exempt.

7. **Errors are typed.** Every raised library-level error subclasses
   `DocxPlusError` (defined in `core/__init__.py`). Some dual-inherit
   `ValueError` or `TypeError` for callers that still catch the stdlib
   bases. See §7.

8. **No unrequested side effects on the input document.** Functions
   that mutate document state document the mutation in the docstring.
   `resolve_*` and `read_*` functions are pure reads.

---

## §7 Error hierarchy

Every library-raised exception subclasses `DocxPlusError`. A few also
dual-inherit a stdlib base when an existing API contract (or SPEC
sentence) calls for it.

| Exception | Bases | Raised from | Meaning |
|---|---|---|---|
| `DocxPlusError` | `Exception` | `core/__init__.py` | Root of the hierarchy. Catch this to catch every library error |
| `DuplicateIdError` | `DocxPlusError`, `ValueError` | `core/ids.py` | `IdRegistry.reserve(n)` called on an already-issued value |
| `StyleExistsError` | `DocxPlusError` | `styles/modify.py` | `create_style` called on an ID already defined |
| `StyleNotFoundError` | `DocxPlusError` | `styles/modify.py` | `apply_style`/`modify_style`/`delete_style` referenced an undefined ID |
| `StyleInUseError` | `DocxPlusError` | `styles/modify.py` | `delete_style` (without `force=True`) on a referenced style |
| `UnknownStylePropertyError` | `DocxPlusError`, `TypeError` | `styles/modify.py` | Unrecognised `**properties` kwarg. SPEC §5 says these raise `TypeError`; dual inheritance lets both contracts hold |
| `StyleCascadeError` | `DocxPlusError` | `styles/inspect.py` | `basedOn` chain cycles or exceeds depth 11 |
| `MissingPartError` | `DocxPlusError` | `styles/inspect.py` | A referenced part is required but absent (currently unused — see §2 layer 4) |
| `ThemeError` | `DocxPlusError` | `styles/theme.py` | Structurally invalid theme input to the transform functions |

The dual-inheritance pattern (`DuplicateIdError`, `UnknownStylePropertyError`)
exists because SPEC sentences predating §9.7's typed-error invariant
documented `ValueError`/`TypeError` as the raised type. Rather than
breaking the spec contract, both bases sit on the class — `except
ValueError` and `except DocxPlusError` both catch.

---

## §8 Testing strategy

SPEC §10 specifies three layers:

- **Layer 1 — structural unit tests.** One file per module, fast,
  no I/O beyond reading fixtures. Currently 165+ tests covering
  `core/`, `styles/theme`, `styles/inspect`, `styles/modify`.
- **Layer 2 — round-trip tests.** Build → save → reopen with
  `python-docx` → assert. The high-value class for OOXML
  correctness (`IMPLEMENTATION.md §8`). Two tests at end of Phase 3.5;
  `TEST_GAPS.md` I1 lists the missing ones.
- **Layer 3 — headless render smoke.** Run each example, convert to
  PDF with LibreOffice headless, assert exit-0 and page count. Gated
  on the `requires_libreoffice` pytest marker; deferred to Phase 6.

Test fixtures live in `tests/fixtures/build_fixtures.py` (the build
script is the source of truth, not the `.docx` files it produces —
`.gitignore` excludes the generated docx files). `empty.docx`,
`multistyle.docx`, and `themed.docx` exist as of Phase 2.

Shared assertions live in `docx_plus/_testing/ooxml_asserts.py`. The
module is internal — not re-exported from the top-level package — and is
built out lazily as later tests demand more helpers.

For a frozen snapshot of where the suite has real holes, see
[`TEST_GAPS.md`](TEST_GAPS.md).

---

## §9 What's next

Phases 4 (Forms), 5 (Fields + Protection), 6 (Polish: examples,
LibreOffice smoke tests, CI doc generation) are still ahead. The
architectural shape they will fit into is fixed: each is a sibling
under `docx_plus/`, each imports from `core/` only, each enforces the
invariants in §6 at its own boundary. See `IMPLEMENTATION.md` for the
phase-by-phase plan; see `SPEC.md` for the contract.
