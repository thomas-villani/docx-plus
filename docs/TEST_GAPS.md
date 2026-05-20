# Test Suite Gap Audit

> **Status (2026-05-20):** The per-item audit below was snapshotted at
> end-of-Phase-5 (v0.1) and has **not** been re-run against the v0.2 suite.
> Treat its stats as historical: the current suite is **717 tests across 34
> files** (pytest, mypy `--strict`, ruff `check` all green). The v0.1 and
> v0.2 cycles did not formally close the IMPORTANT gaps catalogued here —
> later test growth narrowed several, but they remain the priority backlog
> for a v0.3 re-audit. The original "realistic target before Phase 4"
> milestone was not met; that planning note is left below as a record.

**Snapshot date:** 2026-05-19 (end of Phase 5)
**Scope:** Phases 1–5 — `core/`, `styles/`, `controls/`, `fields/`, `protection/`
**Suite size at snapshot:** 285 tests across 17 files; mypy `--strict`, ruff `check`, ruff `format --check` all green on both `docx_plus/` and `tests/`

This is an honest accounting of where the existing test suite has real holes,
not a comprehensive coverage report. Items are listed by severity. Each entry
cites the file/line that exercises (or fails to exercise) the path and the
SPEC / `IMPLEMENTATION.md` section that motivates it. Phase 6 (`examples/`)
is intentionally excluded — it is not yet built, so missing tests there are
expected and not a "gap" in the sense used here.

Entries close (move to a `## Resolved` section at the bottom) as gaps are
filled. Update this file when you fix one.

---

## BLOCKER

_(none — see ## Resolved for B1.)_

---

## IMPORTANT

Correctness gaps that could bite under realistic workloads. Listed in
recommended fix order.

### I1. Save→reopen round-trips cover only 2 of 5 modify operations

- **Where:** `tests/test_styles_modify.py:884` (`test_save_reopen_preserves_created_style`)
  and `tests/test_styles_modify.py:911` (`test_save_reopen_preserves_ensure_style_heading`)
  are the only tests that go to disk and reopen with `python-docx`. The other
  three operations (`modify_style`, `apply_style`, `delete_style`) are tested
  in-memory only.
- **Plausible bug:** A modification that works against the in-memory lxml
  tree but produces XML that `python-docx`'s serializer or Word's parser
  rejects would not surface. Schema-order violations on `modify_style` (see
  I2) are the canonical example.
- **SPEC / IMPL ref:** SPEC §10 Layer 2 ("round-trip tests"),
  `IMPLEMENTATION.md §8` ranks round-trips as the highest-value test class.
- **Fix:** Add three round-trip tests:
  1. `test_save_reopen_preserves_modified_style` — create a style with
     multiple properties, save, reopen, `modify_style` to change two, save,
     reopen, assert resolved formatting reflects both the original
     untouched properties and the modified ones
  2. `test_save_reopen_preserves_applied_style` — create + apply, save,
     reopen, resolve the paragraph, assert it picks up the style
  3. `test_save_reopen_delete_force_leaves_dangling_refs` — apply a style
     to a paragraph, `delete_style(force=True)`, save, reopen, assert the
     `w:pStyle` body reference still names the deleted style (closes I5
     simultaneously)

### I2. Schema-order assertions exist for `create_style` but not `modify_style`

- **Where:** `tests/test_styles_modify.py:277` (`test_style_children_ordered_correctly`),
  `:301` (`test_ppr_children_ordered_correctly`), and `:320`
  (`test_rpr_children_ordered_correctly`) all verify schema order after
  `create_style`. No equivalent test runs after `modify_style`.
- **Plausible bug:** `modify_style` with kwargs that target the same
  composite element (`w:ind`, `w:spacing`, `w:rFonts`) uses merge semantics
  — re-reading the element, mutating attributes, and writing it back. If
  the merge path appends instead of `_ordered_insert`-ing, child order can
  drift. The result remains structurally well-formed XML, so in-memory
  tests pass, but Word silently "repairs" the file on open.
- **SPEC / IMPL ref:** `IMPLEMENTATION.md §4` flags silent repair as a
  high-risk failure mode.
- **Fix:** Add `test_modify_style_preserves_pPr_child_order` and
  `test_modify_style_preserves_rPr_child_order`. Construct a style with
  child elements deliberately near both ends of the canonical order, run
  `modify_style` to touch an attribute on a child mid-list, and assert the
  child order matches `_PPR_CHILD_ORDER` / `_RPR_CHILD_ORDER`.

### I3. Cycle detection only exercised with 2-node cycles

- **Where:** `tests/test_styles_inspect.py` has one cycle test (A→B→A) and
  one depth-limit test (13-node chain triggering `_MAX_STYLE_CHAIN_DEPTH`
  at `inspect.py:33`).
- **Plausible bug:** A defect in the visited-set logic in
  `_collect_style_chain` (`inspect.py:376-399`) could pass the 2-node case
  but mishandle longer cycles or self-references. Three nodes is the
  smallest cycle that distinguishes "checks predecessor" from "checks set
  membership"; a self-loop (A → basedOn A) is the smallest case that
  tests whether `current_id in visited` runs before the chain is appended
  to.
- **SPEC / IMPL ref:** SPEC §4 ("Cycle detection required"),
  `IMPLEMENTATION.md §5` ("Cycle detection lives in one place").
- **Fix:** Add `test_three_node_cycle_raises_with_full_path` (A→B→C→A,
  asserting the error message includes all three IDs in the cycle path)
  and `test_self_referential_basedon_raises` (A → basedOn A).

### I4. Toggle parity not exercised end-to-end with `w:val="0"`

- **Where:** `tests/test_cascade_toggles.py:172`
  (`test_value_zero_treated_as_explicit_false`) confirms the toggle helper
  treats `"0"` the same as `"false"` at the unit level, but no test runs
  the full cascade walker (`_Accumulator.toggle` at `inspect.py:223-238`)
  against a style chain where one layer emits `w:val="0"` and another
  emits unqualified `<w:b/>`.
- **Plausible bug:** Word can emit either `"0"` or `"false"` depending on
  version and authoring tool. A regression that compares `val_attr ==
  "false"` instead of `val_attr in ("0", "false")` would let one of those
  XOR through the chain when it should reset.
- **SPEC / IMPL ref:** SPEC §4 ("Implementation note: track each toggle
  property's parity..."), `IMPLEMENTATION.md §5` toggle test list (case
  3: "Style A bold, style B basedOn A `w:b w:val='false'` → effective
  false").
- **Fix:** Parametrize the existing toggle-chain tests with both `"0"`
  and `"false"` forms of the explicit-reset case, or add a dedicated
  `test_toggle_val_zero_resets_parity_through_basedon_chain`.

### I5. `delete_style(force=True)` only smoke-checked

- **Where:** `tests/test_styles_modify.py:497`
  (`test_delete_style_force_overrides_reference_check`) asserts the style
  is gone after `force=True`, but does not check the body refs.
- **Plausible bug:** A defensive "rewrite to Normal" branch could be
  added (e.g. while refactoring) and the test would still pass — but the
  documented contract (SPEC §5: "leaves dangling references — Word will
  fall back to Normal") would silently change.
- **SPEC ref:** SPEC §5 (`delete_style`).
- **Fix:** Convert the existing assertion to also XPath the body for
  `w:pStyle[@w:val='ToDelete']` and assert at least one such reference
  remains. Folds into I1's `test_save_reopen_delete_force_leaves_dangling_refs`.

### I6. Built-in style materialization is one bulk smoke test

- **Where:** `tests/test_styles_modify.py:609`
  (`test_ensure_style_all_known_builtins_succeed`) iterates every entry
  in `_BUILTIN_STYLES` (`modify.py:1154`, **107 entries** across seven
  tiers) and asserts `ensure_style` returns a proxy. Heading1 has a
  dedicated round-trip at `:584` and `:911`, but no other built-in has
  property-level round-trip coverage — they're only checked for "doesn't
  raise."
- **Plausible bug:** A missing or wrong property in one built-in's
  `properties` dict (e.g. Heading3 missing `spacing_before`,
  `TableofFigures` keyed under the wrong case as happened on
  2026-05-19) materializes successfully — the `properties=` kwargs are
  optional and `ensure_style` falls through to `create_style` on a
  table miss — but the resolved paragraph using that style lacks the
  expected formatting. The bulk test never inspects the resolved output.
- **SPEC ref:** SPEC §5 ("Test that materialization of each produces a
  style Word accepts.")
- **Fix:** Parametrize a single test over a list of `(style_id,
  expected_properties)` tuples; for each, `ensure_style`, apply to a
  paragraph, `resolve_effective_formatting`, and assert each
  `expected_properties` field matches. With 107 entries in the table now
  this is high-value: a regression like the `TableOfFigures`→
  `TableofFigures` styleId case bug would be caught immediately. Start
  with the 14 SPEC §5 "at minimum" set, then expand to all
  sample-sourced tiers (E/F/G).

### I7. `remap_styles` four-step resolution not tested as a sequence

- **Where:** `tests/test_styles_modify.py:724-840` covers each of the four
  steps individually (exact match `:724`, explicit mapping `:732`,
  matcher `:749`, create-from-builtins `:782`), but no test triggers
  fall-through across all four on a single target.
- **Plausible bug:** A step-ordering defect that skips a step entirely
  (e.g. fall-through condition inverted) would not surface — each step's
  test sets up the world so that its step is the one that fires.
- **IMPL ref:** `IMPLEMENTATION.md §12` (Phase 3.5 entry) describes the
  intended four-step chain.
- **Fix:** Add `test_remap_styles_falls_through_all_four_steps` —
  construct a doc with one target that needs each step in turn (across
  multiple targets in one call), assert each lands at the expected step.

### I8. Unknown theme-name graceful-degradation path untested

- **Where:** `styles/theme.py` `ThemeColors.base()` returns `None` for
  names not in `_THEME_NAME_TO_SCHEME_KEY`. `inspect.py:605-620`
  (`_resolve_color`) then sets `acc.partial = True` and returns the
  unresolved name. `tests/test_theme_edge_cases.py` covers strip-rel and
  corrupt-blob scenarios but not the "themeColor is a garbage name"
  path through the cascade walker.
- **Plausible bug:** A future refactor that raises instead of returning
  `None` would turn the inspector into a function that crashes on
  diverse real-world inputs — `IMPLEMENTATION.md §5` explicitly warns
  against this.
- **SPEC ref:** SPEC §4 ("If the theme part is missing or malformed,
  set `partial=True`...").
- **Fix:** Add a test that builds a style with `themeColor="garbage"`,
  resolves a paragraph using it, asserts `partial=True` and
  `color_rgb == "garbage"`.

---

## NICE-TO-HAVE

Defensive completeness — worth adding but not blocking Phase 4.

### N1. Shared assertion library — two helpers still missing

- **Where:** `docx_plus/_testing/ooxml_asserts.py` now exports
  `assert_ids_unique`, `assert_style_defined`, `count_controls`,
  `assert_protected`, `assert_field_dirty`. SPEC §10 also lists
  `assert_style_not_defined` and `assert_no_orphan_relationships`,
  neither of which exists yet.
- **Status:** Reduced from "two-of-seven" to "five-of-seven" across
  Phases 4 and 5. `assert_style_not_defined` has obvious callers in
  current code (the delete-style tests use ad-hoc XPath instead);
  `assert_no_orphan_relationships` is still blocked on a real caller
  needing it.
- **Fix:** Add `assert_style_not_defined`. Defer the relationship
  helper until v0.2 binding work needs it.

### N2. ~~Conditional table formatting has no skip-marked placeholder~~ — **resolved** (v0.2 expansion)

See ## Resolved below — `tests/test_styles_table_conditional.py`
exercises the live implementation rather than carrying an xfail.

---

## Correction note

A draft of this audit flagged `apply_lum_mod` / `apply_lum_off` as
untested. They are actually exercised at `tests/test_styles_theme.py:103,
108, 112, 116` (4 tests covering identity, halving, clamp-at-one,
floor-at-zero). The entry was removed before publication.

---

## Recommended priority order for resolution

Filed in the order that maximizes value-per-effort and unblocks Phase 4:

1. **B1** (coverage threshold) — single PR, single config change, biggest
   structural improvement
2. **I1 + I5** (save→reopen for modify/apply/delete, folded together)
3. **I2** (schema-order on `modify_style` merge path)
4. **I3** (3-node and self-referential cycles)
5. **I6** (per-style cascade verification across built-ins)
6. **I4 + I7 + I8** (parametric / fall-through / theme-unknown edge cases)
7. **N1** (`assert_style_not_defined`)
8. **N2** (xfail placeholder for conditional table formatting)

Items 1–4 were the realistic target before Phase 4 began; 5–8 to land
alongside Phase 4 work. (Historical: this milestone was not met — see the
status note at the top of the file. Re-prioritise during the v0.3 audit.)

---

## Resolved

### 2026-05-19 — v0.2 in-place expansion

- **N2 — Conditional table formatting now lands real cascade coverage.**
  The cascade resolver applies `<w:tblStylePr>` branches
  (`firstRow`, `lastRow`, `firstCol`, `lastCol`, `band1Horz`,
  `band1Vert`, the four corners, `wholeTable`) in ECMA-376 17.7.6.5
  precedence order. 13 tests in `tests/test_styles_table_conditional.py`
  verify auto-derivation from `_Cell` position, manual `TableContext`
  override, precedence (`firstRow` over `band1Horz`, `nwCell` over
  `firstRow`), `wholeTable` always-applies, and that paragraphs / runs
  inside cells inherit the conditional formatting. The xfail
  placeholder the original entry recommended was never needed — the
  feature shipped instead.

### 2026-05-19 — Phase 6

- **N3 — Headless LibreOffice smoke tests shipped.** The
  `requires_libreoffice` marker is now used by
  `tests/test_examples_libreoffice.py` (one test per example script:
  builds, converts to PDF via headless `soffice`, asserts conversion
  succeeds and the PDF has the expected page count). Skipped
  automatically when `soffice` is not on PATH; the Ubuntu/Python 3.13
  CI job installs LibreOffice to exercise them.
- **Coverage scope note.** The 90% gate runs over `core/`, `styles/`,
  and `controls/`. `docx_plus/_testing/*` (test infrastructure) and
  `docx_plus/examples/*` (covered by Layer 3 smoke tests, not unit
  tests) are listed in `pyproject.toml`'s `[tool.coverage.run] omit`
  by design.
- **B1 — Coverage gate flipped on.** `fail_under = 90` added to
  `[tool.coverage.report]` in `pyproject.toml`; `ci.yml` now runs
  `pytest ... --cov-fail-under=90`. Coverage at 91.76% aggregate
  (styles 90.7%, core 91%, controls 93%). Bringing this from a stale
  77% on `styles/inspect.py` required two new test files:
  `tests/test_cascade_numbering.py` (7 tests covering layer 4 of the
  cascade) and `tests/test_cascade_run_target.py` (5 tests covering Run
  and `_Cell` target paths). The `numbered.docx` fixture in
  `tests/fixtures/build_fixtures.py:build_numbered` materialises a
  custom `abstractNum` / `num` pair inside python-docx's already-shipped
  numbering part.

### 2026-05-19 — Phase 5

- **Partial: N1** — `count_controls` (Phase 4), `assert_protected`
  (Phase 5), `assert_field_dirty` (Phase 5) added to
  `docx_plus/_testing/ooxml_asserts.py`. Five-of-seven SPEC §10 helpers
  now present; entry above is rewritten with the remaining two.

## Phase 4–5 coverage notes

These don't belong in the BLOCKER/IMPORTANT/NICE-TO-HAVE buckets above
— they're scoped to features that landed in Phases 4 and 5 — but they
record what is and isn't tested for those modules.

- **`controls/builder.py`** — 24 tests in `tests/test_controls_builder.py`
  cover every control type's structure assertion + a save→reopen
  round-trip, multi-builder ID-registry sharing, the
  `MissingNamespaceError` path, and sdtPr child-order invariants per
  control type.
- **`controls/read.py`** — 26 tests in `tests/test_controls_read.py`
  cover the four typed errors, every `set_control_value` type path,
  the dropdown-by-value vs by-display matching, the combobox
  freeform-passthrough, and the `existing_form` fixture (a doc built
  without `FormBuilder` — schema-tolerance check).
- **`fields/`** — 24 tests in `tests/test_fields.py`. Every field
  variant's instruction text, the `xml:space="preserve"` invariant,
  two-fields-in-one-paragraph composition, three save→reopen
  round-trips (PAGE, DATE, mark-dirty+PAGE), and four
  `mark_fields_dirty` cases (insert / idempotent / update-false /
  schema-position / fallback-append).
- **`protection/`** — 18 tests in `tests/test_protection.py`. Every
  `ProtectionMode` literal (parametrised), the schema-position
  invariant against `w:defaultTabStop`, idempotency + mode-replace,
  unprotect+reprotect, the `is_protected` predicate, save→reopen
  round-trip, plus negative tests on the `assert_protected` helper.
- **Round-trip floor for new capabilities**: every modifying function
  in `controls/`, `fields/`, `protection/` has at least one
  save→reopen test. The Phase 3 gap on `modify_style` / `apply_style` /
  `delete_style` (I1, I5 above) is the outlier, not the rule.

_Move entries here with their resolution date and the test file/line
that now covers them._
