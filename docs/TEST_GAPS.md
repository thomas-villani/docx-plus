# Test Suite Gap Audit

**Snapshot date:** 2026-05-15 (end of Phase 3.5)
**Scope:** Phases 1, 2, 3, 3.5 — `core/`, `styles/inspect.py`, `styles/modify.py`, `styles/theme.py`
**Suite size at snapshot:** 187 tests across 12 files; mypy `--strict`, ruff `check`, ruff `format --check` all green

This is an honest accounting of where the existing test suite has real holes,
not a comprehensive coverage report. Items are listed by severity. Each entry
cites the file/line that exercises (or fails to exercise) the path and the
SPEC / `IMPLEMENTATION.md` section that motivates it. Phase 4–6 capabilities
(`controls/`, `fields/`, `protection/`, `examples/`) are intentionally
excluded — they are not yet built, so missing tests there are expected and
not a "gap" in the sense used here.

Entries close (move to a `## Resolved` section at the bottom) as gaps are
filled. Update this file when you fix one.

---

## BLOCKER

Gaps that could let a quality-gate violation or silent document-corruption
bug ship.

### B1. Coverage threshold defined in SPEC but not enforced

- **Where:** `pyproject.toml:92-104` defines `[tool.coverage.run]` and
  `[tool.coverage.report]` but neither sets `fail_under`, and no CI step
  passes `--cov-fail-under=90`.
- **SPEC ref:** §13 mandates "Coverage ≥ 90% on `core/`, `styles/`,
  `controls/`" as a hard quality gate.
- **Risk:** As the surface grows in Phases 4–6, untested branches will
  accumulate with no signal. The gate exists in writing only.
- **Fix:**
  1. Add `fail_under = 90` to `[tool.coverage.report]` in `pyproject.toml`
  2. In `.github/workflows/ci.yml`, replace `pytest` with `pytest --cov=docx_plus --cov-fail-under=90` (or equivalent)
  3. Run locally first to confirm the current suite clears the bar; if
     not, the failing lines are themselves gaps to fill before flipping
     the gate on

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

### N1. Shared assertion library is sparse

- **Where:** `docx_plus/_testing/ooxml_asserts.py` exports
  `assert_ids_unique` and `assert_style_defined`. SPEC §10 also lists
  `assert_style_not_defined`, `assert_no_orphan_relationships`,
  `assert_protected`, `assert_field_dirty`, `count_controls`.
- **Status:** Four of the five missing helpers (`assert_protected`,
  `assert_field_dirty`, `count_controls`, `assert_no_orphan_relationships`)
  are legitimately blocked on Phases 4–5 — no callers yet. The fifth,
  `assert_style_not_defined`, has obvious callers in current code (e.g.
  the delete-style tests use ad-hoc XPath instead).
- **Fix:** Add `assert_style_not_defined`. Defer the others until Phase 4
  starts.

### N2. Conditional table formatting has no skip-marked placeholder

- **Where:** `styles/inspect.py:402-422` (`_apply_table_style_chain`)
  documents the deferral inline, but no test file records the intent.
- **Risk:** A future contributor might assume conditional formatting works
  and write a test against it that passes for the wrong reason (base style
  matches the conditional style coincidentally).
- **Fix:** Add `tests/test_cascade_table.py` with a single
  `@pytest.mark.xfail(reason="conditional w:tblStylePr deferred per SPEC §4 step 2", strict=True)`
  test that constructs a `w:tblStylePr` and asserts it influences the
  cascade. xfail-strict means the test will start failing the moment
  someone fixes the underlying behavior, which is the desired signal.

### N3. No headless LibreOffice smoke test exists yet

- **Where:** `pyproject.toml:87-89` declares the
  `requires_libreoffice` pytest marker but no test uses it.
- **SPEC ref:** SPEC §10 Layer 3 explicitly defers this to the polish
  phase; this entry is just to record that it's tracked.
- **Fix:** Phase 6.

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

Items 1–4 are the realistic target before Phase 4 begins; 5–8 can land
alongside Phase 4 work.

---

## Resolved

_None yet. Move entries here with their resolution date and the test
file/line that now covers them._
