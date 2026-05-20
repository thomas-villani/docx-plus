# docx-plus v0.2.0 — Pre-Release Code & Docs Review

Synthesised from five parallel review passes against the v0.2 in-place
expansion (commits `ae2abbc..HEAD`). Findings are grouped by severity,
each tagged with the subsystem so unrelated issues can be resolved in
separate sessions without context bleed.

Each finding has a stable id (`C1`, `H4`, `M11`, etc.) — use it in
commit messages and PR titles so we can cross-reference here.

## Status

- **Session A — cascade correctness — DONE.** C2, H1, H2, H4, H5
  resolved (11 new tests; mypy strict + ruff clean). H3 closed as
  **wontfix** — see annotation under the finding.
- **Session B — schema / part wiring — DONE.** C1, C3, H6, H7, H8
  resolved (5 new tests; mypy strict + ruff clean).
- **Session C — error taxonomy + interleaving — DONE.** C4, H9, H10,
  M1, M2 resolved (7 new tests; mypy strict + ruff clean).
- **Session D — publishing API hardening — DONE.** H11, H12, H13,
  M14, M15, M16 resolved (37 new tests; 616 total pass; mypy strict +
  ruff clean).
- Sessions E–F pending.

## Stats

| Severity   | Count |
| ---------- | ----- |
| Critical   | 5     |
| High       | 17    |
| Medium     | 23    |
| Low        | 21    |
| Nit        | 13    |
| **Total**  | **79**|

## Triage notes / recommended session order

The findings cluster naturally into six work sessions; each could land
as its own PR. Suggested order is correctness-first, then docs, then
polish:

1. **Session A — cascade correctness** (Critical C2, High H1, H2, H3, H4, H5)
   The styles cascade has the highest concentration of real-bug-with-no-test findings.
   This is where round-trip fidelity is most at risk.
2. **Session B — schema / part wiring** (Critical C1, C3, High H6, H7, H8)
   Missing footnote separators and the `pgBorders` child-order bug
   produce files some strict consumers (and some Word versions) will
   reject or "repair".
3. **Session C — error taxonomy & SPEC alignment** (Critical C4, High H9, H10, M1, M2)
   SPEC §9.7 vs ARCHITECTURE §9 currently contradict each other; the
   error table is missing four shipped classes.
4. **Session D — publishing API hardening** (High H11, H12, H13, M3, M4, M5, M6, M7)
   Field-instruction injection + validation gaps; small DX improvements.
5. **Session E — docs / release-day** (Critical C5, High H14, H15, H16, H17, M8–M19, plus all Nit)
   Stale dates, stale "deferred to v0.2" docstrings, mkdocs nav gaps,
   release-day decisions (PyPI banner, alpha→beta classifier).
6. **Session F — tests & smells** (Low L1–L21, plus M20–M23)
   Brittle assertions, coverage gaps, the conftest fixture-build duality
   the user already flagged in `notes.md`.

A pre-tag "must-fix" cut would be C1–C5 + H1, H2, H4, H6, H9 — eight
items. Everything else can ship in v0.2.1 if needed.

---

## Critical

### C1: Footnotes/endnotes parts created without required separator entries
- **Status:** ✅ RESOLVED (Session B) — `FOOTNOTES_SPEC` and `ENDNOTES_SPEC` now use a new `_notes_root_with_separators` helper that pre-seeds the part with the two reserved entries on first creation. `read_footnotes` / `read_endnotes` already filter ids ≤ 0 so user-visible iteration is unaffected. Round-trip tests pin both the seed presence and the `w:type` attribute on each separator.
- **Subsystem:** core / notes
- **Location:** `docx_plus/core/parts.py:95-99` (`_empty_root`), `docx_plus/notes/write.py:_add_note`
- **Description:** When `add_footnote` / `add_endnote` first runs, `get_or_create_part` fabricates a new `footnotes.xml` / `endnotes.xml` from `_empty_root("footnotes")` which produces just `<w:footnotes/>` — no `<w:footnote w:id="-1" w:type="separator">` or `<w:footnote w:id="0" w:type="continuationSeparator">` entries. Word expects those entries to be present and uses them to render the horizontal separator line between body text and the first footnote on a page; some Word versions will offer a "repair file" prompt when opening a docx that has user-authored footnotes but no separator entries. This is the primary reason the registry code carefully reserves ids `-1`/`0` — they're meant to be in the part. The library never writes them.
- **Suggestion:** Have `_add_note` (or `get_or_create_part` via a per-spec "initial content" hook) emit the two separator footnotes (and equivalent endnote entries) the first time a notes part is created. Each separator is `<w:footnote w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` and `<w:footnote w:id="0" w:type="continuationSeparator"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>`; endnotes use the same shape under `<w:endnote>` with `<w:endnoteRef>` style.

### C2: Run-level `rStyle` applied AFTER direct run formatting (wrong precedence)
- **Status:** ✅ RESOLVED (Session A) — order swapped in `_apply_paragraph_cascade`; new `runStyle` Layer literal added; regression test `test_run_rstyle_is_overridden_by_direct_run_rpr` pins the fix.
- **Subsystem:** styles cascade
- **Location:** `docx_plus/styles/inspect.py:407-414`
- **Description:** In `_apply_paragraph_cascade` the run cascade does, in order: linkedCharStyle → directRun → run's own `rStyle` chain. But ECMA-376 17.3.2.29 makes `rStyle` a *style reference* whose layer sits **below** direct run formatting (analogously to how `pStyle` sits below direct paragraph formatting in this very file at lines 369-398). The current order means a run with `<w:rStyle w:val="MyChar"/>` AND a direct `<w:b/>` will end up XOR-toggled-off if `MyChar` also defines bold, and any non-toggle property on `MyChar` will *overwrite* the direct run formatting — the opposite of what Word does. `test_resolve_run_target_with_rstyle_applies_linked_char_chain` only checks rStyle in isolation, so this is invisible to the suite. The bug is also visible in provenance: the `linkedCharStyle` layer is reused for run-level `rStyle`, conflating two different cascade roles.
- **Suggestion:** Move the `run_style_id` block (lines 411-414) so it runs *before* the `run_rpr` block at 408-410, and promote it to a distinct layer name (e.g. `"runStyle"`) instead of reusing `"linkedCharStyle"` so provenance can tell the two apart. Add a regression test: rStyle setting bold + direct `<w:b w:val="false"/>` on the same run must resolve `bold=False`.

### C3: `set_page_borders` emits children in `top,bottom,left,right` order — schema requires `top,left,bottom,right`
- **Status:** ✅ RESOLVED (Session B) — emission loop reordered to schema sequence. New `test_set_page_borders_child_order_is_top_left_bottom_right` walks the children and asserts the exact order (not just set membership).
- **Subsystem:** layout
- **Location:** `docx_plus/layout/borders.py:110-115`
- **Description:** ECMA-376 17.6.10 `CT_PageBorders` defines the four side children as a fixed sequence `top → left → bottom → right`. The current code loops in the order `top, bottom, left, right`. A schema-strict consumer (Word's xml validator in some configurations; downstream OOXML libraries; ECMA-376 schema validators) will reject the output. python-docx's own internal validators are permissive so the local round-trip tests pass, masking the bug.
- **Suggestion:** Change the loop tuple to `(("w:top", top), ("w:left", left), ("w:bottom", bottom), ("w:right", right))`. Add a test that walks `list(pg)` and asserts the child tag order matches schema order, not just set membership.

### C4: SPEC §9.7 / §16 contradicted by v0.2 raw `ValueError` raises (governance)
- **Status:** ✅ RESOLVED (Session C, doc-only) — SPEC §9.7 rewritten as "typed for domain conditions" with explicit permission for raw `ValueError`/`TypeError` at argument-shape boundaries. SPEC §16 gained a "Raw-exception carve-out (v0.2+)" subsection enumerating the concrete sites and articulating the "does catching by class help?" dividing line. `CommentNotFoundError` and `NoteNotFoundError` added to the §16 table. Now agrees with ARCHITECTURE §9.
- **Subsystem:** core / governance
- **Location:** `SPEC.md:917-948` vs `docs/ARCHITECTURE.md:791-797` and many capability sites:
  `docx_plus/notes/write.py:215,286`, `docx_plus/bookmarks/anchor.py:101,189`, `docx_plus/comments/anchor.py:298`, `docx_plus/layout/breaks.py:80,86`, `docx_plus/layout/columns.py:67,69`, `docx_plus/layout/line_numbering.py:91,93,95`
- **Description:** SPEC §16 says "Library code never raises raw `ValueError`/`RuntimeError`/`TypeError` for caller-facing conditions" and lists only two carve-outs (`IdRegistry.next` RuntimeError; `TypeError` at the `Paragraph`/`Run`/`_Cell` boundary). Every v0.2 capability raises raw `ValueError` for things outside those carve-outs: bookmark name validation, "only supports the main document body" guards, `set_columns`/`set_line_numbering` range checks, reserved note-id checks. `ARCHITECTURE.md` §9 explicitly blesses the v0.2 deviation, so the two governing docs are in direct conflict.
- **Suggestion:** Pick one. Either (a) introduce typed errors (`InvalidBookmarkNameError`, `UnsupportedTargetError`, `OutOfRangeError`) for each case and update both SPEC and ARCHITECTURE to enumerate them, or (b) amend SPEC §16 to formally bless the broader carve-out ARCHITECTURE already documents. Right now a strict reader of SPEC §9.7 would flag every site above as a violation. Resolve before tag.

### C5: Release-day decisions: `Pre-publication`, alpha classifier, mkdocs missing SPEC
- **Subsystem:** docs / release
- **Location:** `README.md:38-40`, `docs/index.md:8`, `pyproject.toml:11`, `mkdocs.yml:30-65`
- **Description:** Three independent release-blockers bundled because they all need decisions before `git tag v0.2.0`:
  - README + docs/index banner say `Status: Pre-publication — not yet on PyPI` — this will be wrong on tag day. Decide whether v0.2.0 ships to PyPI.
  - `pyproject.toml` still has `Development Status :: 3 - Alpha`; v0.2.0 with 558 tests and 93% coverage understates maturity vs the conventional Beta classifier pre-1.0.
  - `mkdocs.yml` nav doesn't include `SPEC.md` (it lives at the repo root, not `docs/`), yet shipped docs link to "SPEC §15", "SPEC §1 non-goal", etc. throughout — every such link 404s on the published site.
- **Suggestion:** Decide PyPI publication policy and update banners; bump classifier (or document why staying alpha). For SPEC, either move it into `docs/` and add to nav, or rewrite all in-docs "SPEC.md" cross-refs to GitHub blob URLs.

---

## High

### H1: `_TBL_STYLE_PR_ORDER` puts horizontal bands and column branches in wrong order vs ECMA-376
- **Status:** ✅ RESOLVED (Session A) — tuple reordered (`firstRow`, `lastRow` before `firstCol`, `lastCol`); three new tests cover row-vs-col precedence at corners + 1-row / 1-column edge cases.
- **Subsystem:** styles cascade
- **Location:** `docx_plus/styles/inspect.py:75-89`
- **Description:** ECMA-376 17.7.6.5 / §17.4.31 lists application order as `wholeTable` → `band1Vert` → `band2Vert` → `band1Horz` → `band2Horz` → `firstRow` → `lastRow` → `firstCol` → `lastCol` → corners. Critically, **rows before columns** so column branches override row branches at intersections (which is why corner branches even need to exist as a separate layer). The code currently puts firstCol/lastCol before firstRow/lastRow. A cell at (row 0, col 0) with both `firstRow` and `firstCol` branches will resolve firstRow's properties as winning — but in real Word, firstCol wins. `test_corner_overrides_first_row` doesn't catch this because no `firstCol` branch is provided.
- **Suggestion:** Swap the relative order so the tuple reads `…, "firstRow", "lastRow", "firstCol", "lastCol", "nwCell", …`. Add a test where a 3×3 table has only `firstRow` and `firstCol` (no corner) and assert the top-left resolves to firstCol's properties.

### H2: `dstrike` is excluded from the cascade — never read, never clears `strike`
- **Status:** ✅ RESOLVED (Session A) — `dstrike` is read in `_apply_rpr` as a non-toggle property (last-writer-wins per ECMA-376 17.7.3); surfaced as `ResolvedFormatting.double_strike`. Note: the reviewer's "clear `strike` when `dstrike` is set" suggestion was rejected — strike and dstrike are independent per spec; mutual exclusivity is a Word UI convention, not a file-format rule. Four new tests in `test_cascade_toggles.py`.
- **Subsystem:** styles cascade
- **Location:** `docx_plus/styles/inspect.py:107-108`, `_apply_rpr:779-820`
- **Description:** ECMA-376 17.3.2.9 (`strike`) and 17.3.2.10 (`dstrike`) are mutually exclusive — setting one is supposed to clear the other in resolved formatting. The current resolver never reads `dstrike` at all, never surfaces it on `ResolvedFormatting`, and never clears `resolved.strike` when a later layer sets `dstrike`. A user with a base style `strike=True` and a child style adding `dstrike` will incorrectly resolve to `strike=True`. Comment correctly states dstrike is not a toggle; the bug is the resolver dropping it entirely.
- **Suggestion:** Add `dstrike` extraction in `_apply_rpr` (write to a new `double_strike: bool | None` field, or at minimum clear `strike` when `dstrike` is set non-false). Decide whether `ResolvedFormatting` should surface `dstrike` for v0.2 or note as a known limitation.

### H3: `<w:rPr>` inside `<w:pPr>` is silently dropped when target is a Run
- **Status:** ❌ WONTFIX (Session A) — reviewer misread ECMA-376. **17.3.1.31** scopes `pPr/rPr` to "the glyph used to represent the physical location of the paragraph mark" — i.e., the pilcrow itself. In Word, formatting applied only to the paragraph mark does not cascade to existing runs in the paragraph (verifiable: select the ¶ in show-marks mode, apply bold — only the ¶ becomes bold, not the surrounding text). Current behaviour (surface for paragraph targets, skip for run targets) is correct.
- **Subsystem:** styles cascade
- **Location:** `docx_plus/styles/inspect.py:396-398`
- **Description:** Paragraph-mark formatting (`<w:pPr><w:rPr>…`) is the run formatting of the implicit run that terminates the paragraph. Per ECMA-376 17.3.1.29 it applies to runs in the paragraph as a baseline overridable by run-level rPr. The current guard `and run_element is None` means: when resolving a Run target, the paragraph-mark rPr is **not applied at all** — the run sees only its own `<w:rPr>` and the style chain. A paragraph whose `<w:pPr><w:rPr><w:b/></w:rPr></w:pPr>` sets bold should make every run in that paragraph bold (assuming the run doesn't override), but with the current logic a run target reports `bold=None`.
- **Suggestion:** Remove the `and run_element is None` guard so the directParagraph rPr always applies. Verify existing tests still pass; add a test where the pPr's rPr sets bold and the run has no rPr.

### H4: Band-size attribute (`tblStyleRowBandSize`/`ColBandSize`) ignored
- **Status:** ✅ PARTIALLY RESOLVED (Session A) — `_derive_table_context_from_element` now reads `tblStyleRowBandSize` / `tblStyleColBandSize` from the table instance's own `<w:tblPr>` via a new `_read_band_size` helper. Style-chain lookup (band size declared only on the table style, not on the instance) is documented as a known v0.3 limitation in the `TableContext` docstring. Two new tests cover `band_size=2` (rows) and `band_size=3` (cols).
- **Subsystem:** styles cascade
- **Location:** `docx_plus/styles/inspect.py:601-640` (`_derive_table_context_from_element`)
- **Description:** ECMA-376 17.4.79/80 lets a table style declare `<w:tblStyleRowBandSize w:val="2"/>` so bands span two rows each. The "Grid Table" family of built-in styles use a row band size of 1 but many template styles use 2 or 3. Current derivation hard-codes a band size of 1 (the `row_idx % 2 == 1` check). For tables with non-default band sizes the wrong band branch fires (or no branch when one was expected). Also, the band logic is offset from row 0 unconditionally — but per spec, bands start counting *after* `firstRow` when that branch exists.
- **Suggestion:** Read `w:tblPr/w:tblStyleRowBandSize` / `w:tblStyleColBandSize` (default 1) when computing band parity, and offset the band starting row by 1 only when the resolved table style has a `firstRow` conditional. At minimum, document the current limitation in `TableContext` docstring.

### H5: `band2Horz` / `band2Vert` listed in `_TBL_STYLE_PR_ORDER` but never derivable
- **Status:** ✅ RESOLVED (Session A) — `TableContext` gained `is_band2_row` / `is_band2_col` fields (additive — no break for existing manual callers). `_matching_conditional_types` now emits `band2Horz` / `band2Vert` based on the new fields. Two new tests pin band2 reachability across rows and columns.
- **Subsystem:** styles cascade
- **Location:** `docx_plus/styles/inspect.py:75-89`, `_matching_conditional_types:570-598`
- **Description:** Tuple lists `band2Vert` and `band2Horz` but `_matching_conditional_types` only ever returns `band1Vert` / `band1Horz`; there is no `is_band2_row` / `is_band2_col` field on `TableContext`. Per spec, band2 is the **complement** of band1 — every row is either band1 or band2. Real "Grid Table 4 — Accent 1" and similar styles use band1 for the lighter stripe and band2 for the darker. Currently those tables render with only half their stripes.
- **Suggestion:** Either (a) extend `TableContext` with `is_band2_row` / `is_band2_col`, derive as the complement of band1 (taking band-size into account), and emit those branches; or (b) explicitly drop band2 strings from `_TBL_STYLE_PR_ORDER` and document the limitation crisply.

### H6: `edit_comment` / `_edit_note` only remove `<w:p>` children — other block-level descendants leak through
- **Status:** ✅ RESOLVED (Session B) — both helpers now strip every child element (unconditionally) before appending the new paragraph. Element-level attributes (`author`/`date`/`initials` on comments; `w:id`/`w:type` on notes) live on the element itself and are preserved. Two new tests inject a `<w:tbl>` into the body and assert it's removed by the edit call.
- **Subsystem:** comments / notes
- **Location:** `docx_plus/comments/anchor.py:176-178`, `docx_plus/notes/write.py:229-231`
- **Description:** Both edit helpers filter children to `localname == "p"`, removing only paragraph children before appending the new paragraph. ECMA-376 17.13.4.2 (`CT_Comment`) and the footnote/endnote equivalents extend `EG_BlockLevelElts`, which legally includes `<w:tbl>`, `<w:customXml>`, `<w:sdt>`, etc. A comment authored in Word with an embedded table will, after `edit_comment`, end up with the OLD table next to the NEW paragraph — a hybrid body the caller never intended.
- **Suggestion:** Strip all children unconditionally (`for child in list(comment_el): comment_el.remove(child)`), then append the new paragraph. Attributes (`author`, `date`, `initials`) live on the element itself, not on its children, so removal is safe.

### H7: `set_page_borders` omits `w:offsetFrom` — visual output won't match Word UI
- **Status:** ✅ RESOLVED (Session B) — `<w:pgBorders w:offsetFrom="page"/>` is now the default, matching Word's UI emission. New `offset_from: Literal["page", "text"] = "page"` kwarg lets callers pick. New `OffsetFrom` literal re-exported from `docx_plus.layout`. Also corrected the `Border.space` docstring (M5 partial): the unit is **points** (range 0-31), not twips.
- **Subsystem:** layout
- **Location:** `docx_plus/layout/borders.py:109`
- **Description:** ECMA-376 17.6.10 declares optional attributes on `CT_PageBorders`: `w:zOrder`, `w:display`, `w:offsetFrom`. Word's UI default for "Page Border: Box, Whole Document" emits `w:offsetFrom="page"` — offsets measured from page edge (the standard decorative-border look). With no `offsetFrom`, Word's documented default is `"text"`, which measures the gap from body text, producing a tight inner box that visually does not match the UI. `Border.space=24` default the docstring describes as "the value Word's UI emits for Whole document, Box, Default settings" is consistent with `offsetFrom="page"`, so the docstring and the emitted XML disagree.
- **Suggestion:** Either emit `w:offsetFrom="page"` by default on `<w:pgBorders>`, or add `offset_from: Literal["page", "text"] = "page"` keyword on `set_page_borders` and write the attribute.

### H8: `clear_all_comments` is O(N×M)
- **Status:** ✅ RESOLVED (Session B) — rewritten as a single body-walk that removes every range marker / reference regardless of id (two `xpath(body, ".//w:commentRange*")` + one `xpath(body, ".//w:commentReference")`), followed by one walk over `comments.xml` removing every entry. Existing 5 `clear_*` tests pass unchanged.
- **Subsystem:** comments
- **Location:** `docx_plus/comments/anchor.py:182-207`
- **Description:** Function calls `delete_comment(doc, comment_id)` in a loop; each call runs three full-body XPaths plus another against the comments part. For N comments and body size M that's O(N×M). Functionally correct, but the function exists specifically as the bulk operation — exactly where the per-comment scan is wrong.
- **Suggestion:** Walk the body once, removing every `w:commentRangeStart`, `w:commentRangeEnd`, and `w:commentReference` (with enclosing run) in a single pass; then walk `comments.xml` once removing every `<w:comment>`. No need to match by id when clearing all.

### H9: `_apply_table_style_chain` walks the basedOn chain twice with wrong interleaving
- **Status:** ✅ RESOLVED (Session C) — rewritten as a single ancestors-first walk that, per style level, applies base pPr/rPr then matching `<w:tblStylePr>` branches in `_TBL_STYLE_PR_ORDER`. `_apply_conditional_table_formatting` folded into the main helper. New regression test pins child-base-overrides-parent-conditional behaviour.
- **Subsystem:** styles cascade
- **Location:** `docx_plus/styles/inspect.py:501-529`
- **Description:** Line 527 calls `_apply_style_chain` (walks chain applying base pPr/rPr). Then if `table_context` is provided, line 529 calls `_apply_conditional_table_formatting` which re-collects the chain (line 547) and walks it again. Beyond the double-work, the *order* doesn't match ECMA-376 17.7.6.5: spec says for each style level (root→leaf), apply base pPr/rPr → matching conditional branches in precedence order. Current code does all base pPr/rPr for the whole chain root-to-leaf, then all conditional branches. A leaf style's base properties override an ancestor's conditional branches — reversed from spec.
- **Suggestion:** Refactor to a single walk that, per style in the chain (ancestors first), applies that style's base pPr/rPr then its matching conditional `tblStylePr` branches in `_TBL_STYLE_PR_ORDER`.

### H10: Duplicated `_insert_before_first_anchor` in `protection/document.py`
- **Status:** ✅ RESOLVED (Session C) — local copy deleted; module now imports `insert_before_first_anchor` from `core.oxml`. Removed dead `lxml.etree` import. All 18 protection tests pass.
- **Subsystem:** core / protection
- **Location:** `docx_plus/protection/document.py:51-65`
- **Description:** Byte-for-byte copy of `docx_plus.core.oxml.insert_before_first_anchor` (right down to the docstring sentence). Predates the shared helper, but is exactly the "reinvented equivalent" v0.2 was supposed to consolidate. Every other settings-touching module (`fields/update.py`, `layout/settings.py`, `layout/borders.py`, `layout/line_numbering.py`) uses the shared one. Drift risk: if the shared helper's semantics change, protection is silently left on the old behaviour.
- **Suggestion:** Replace local function with `from docx_plus.core.oxml import insert_before_first_anchor`; drop the private copy.

### H11: Field-instruction injection via unsanitised string interpolation
- **Status:** ✅ RESOLVED (Session D) — new `docx_plus/publishing/_validate.py` holds shared validators. `add_caption` and `add_table_of_figures` validate `caption_type` against the SEQ identifier rule; `add_caption` validates `numbering` against the ECMA-376 17.16.4.1 token frozenset. All inject-vector inputs (`'Figure" \o "1-9'`, `"Figure,evil"`, empties, leading digits) now raise `ValueError` at function entry. Closes M16 in the same pass.
- **Subsystem:** publishing
- **Location:** `docx_plus/publishing/captions.py:73`, `docx_plus/publishing/figures.py:59`, `docx_plus/publishing/toc.py:50`
- **Description:** Every public helper interpolates user input directly into the OOXML field-instruction string with zero validation or escaping:
  - `add_caption(..., caption_type='Figure" \\f "evil', numbering='ARABIC \\* ROMAN')` produces `SEQ Figure" \f "evil \* ARABIC \* ROMAN`. Per ECMA-376 17.16.5.56 the SEQ identifier must be a single token; the unescaped quote/backslash terminates the identifier and injects additional switches.
  - `add_table_of_figures(caption_type='Figure" \\o "1-9')` produces `TOC \c "Figure" \o "1-9" \h \z` — a ToF that now also picks up paragraphs by outline level, mixing the two TOC modes ECMA-376 17.16.5.61 keeps disjoint.
  Not a security boundary (input is author-controlled), but a correctness bomb for callers who concatenate variable input.
- **Suggestion:** Validate eagerly. `caption_type` must match Word's SEQ-identifier rule (`[A-Za-z][A-Za-z0-9_]*`); `numbering` must be in a frozen set of `\*` format-picture tokens (`ARABIC`, `ROMAN`, `Roman`, `ALPHABETIC`, `alphabetic`, `CardText`, `DollarText`, `Hex`, `Ordinal`, `OrdText`). Raise `ValueError` (resolve with C4) on mismatch.

### H12: No validation on `add_toc(levels=…)` — silently produces broken TOCs
- **Status:** ✅ RESOLVED (Session D) — `validate_outline_levels` rejects non-tuples, wrong arity, non-ints, bools, out-of-range and reversed ranges. 9 parametrized bad-input cases in tests.
- **Subsystem:** publishing
- **Location:** `docx_plus/publishing/toc.py:62-63`
- **Description:** `(3, 1)` (reversed), `(0, 5)` (zero), `(1, 10)` (Word only has 9 levels), `(-1, 3)`, `(1,)`, `()`, or a bare int all flow through `lo, hi = levels` unchecked. `(3, 1)` yields `\o "3-1"` which Word treats as an empty range; `(1,)` raises `ValueError: not enough values to unpack` deep inside the helper with no context. None surface the user-facing problem.
- **Suggestion:** Validate at entry: tuple of length 2, both ints, `1 <= lo <= hi <= 9` (Word's outline-level domain). Raise a typed `ValueError`/`DocxPlusError` naming the bad value.

### H13: `omit_styles` / `\t` switch promised in plan but not implemented
- **Status:** ✅ RESOLVED (Session D) — implemented as `additional_styles: Sequence[tuple[str, int]] | None = None` (renamed from the plan's misleading `omit_styles` — the `\t` switch *adds* styles to the TOC, doesn't omit). Emits `\t "Style1,1,Style2,2"`. Validation in `validate_additional_styles` rejects bad pairs (including style names with `,` or `"` that would terminate the switch).
- **Subsystem:** publishing
- **Location:** `docx_plus/publishing/toc.py:23-29`
- **Description:** The v0.2 expansion plan listed an `omit_styles=None` parameter for `add_toc` plumbed to the `\t` switch (ECMA-376 17.16.5.61: `\t "<style>,<level>,…"` lets the TOC pull from arbitrary style/level pairs instead of the implicit `Heading1..N` set). The shipped implementation has no such parameter.
- **Suggestion:** Either add `omit_styles: Sequence[tuple[str, int]] | None = None` plumbing (and `\t` emission) or update the planning notes and changelog to mark explicitly as deferred to v0.3.

### H14: Test count in `ARCHITECTURE.md` §10 is stale
- **Subsystem:** docs
- **Location:** `docs/ARCHITECTURE.md:824` ("**532 tests** at end of the v0.2 in-place expansion")
- **Description:** Actual test count is 558. The breakdown in §10 ("v0.1's surface (319 tests) plus the v0.2 cycle… 532") no longer matches. `IMPLEMENTATION.md` §12's 2026-05-19 entry also locks in 532.
- **Suggestion:** Update §10 to "558 tests" and refresh the breakdown; update IMPLEMENTATION.md §12 in the same pass.

### H15: Four exported exception classes are undocumented (`IdRangeError`, `InvalidNamespaceError`, `InvalidColorError`, `InvalidDropdownItemError`)
- **Subsystem:** docs / core / styles / controls
- **Location:** `docx_plus/core/__init__.py:23`, `docx_plus/styles/__init__.py:12`, `docx_plus/controls/__init__.py:6` vs `docs/API.md` and `docs/ARCHITECTURE.md` §9
- **Description:** All four are real public exception classes in `__all__`, but they are missing from `docs/API.md` exception tables, `docs/ARCHITECTURE.md` §9's error hierarchy table, and CHANGELOG. `IdRangeError` does get one prose mention in ARCHITECTURE §9 but is missing from the table at line 768.
- **Suggestion:** Add four rows to `docs/ARCHITECTURE.md` §9 and to `docs/API.md` under the appropriate modules. Audit each `docs/reference/*.md` `members:` list against its module's `__all__` to catch this class of drift.

### H16: SPEC.md §15 deferred list + §16 error table are stale relative to current code
- **Subsystem:** docs / SPEC
- **Location:** `SPEC.md:894-948`
- **Description:** §15 is the original v0.1→v0.2 deferred list — mixes "shipped" and "still-deferred" indiscriminately. §16's error table lists 17 errors but is missing v0.2 additions `CommentNotFoundError` and `NoteNotFoundError`. `ARCHITECTURE.md` §9/§11 have corrected lists; SPEC and ARCHITECTURE now disagree on what's shipped.
- **Suggestion:** Either freeze SPEC as a v0.1 contract with a banner ("As of v0.1 spec; see ARCHITECTURE.md §11 for current state"), or rewrite §15 to list only items still deferred at v0.2.0 and add the missing errors to §16. Pair with C4.

### H17: `_TOGGLE_PROPS` in `modify.py` missing six new toggle properties (API asymmetry)
- **Subsystem:** styles modify
- **Location:** `docx_plus/styles/modify.py:341-348` and `_RUN_LEVEL_PROPS:324-339`
- **Description:** `_TOGGLE_PROPS` still only lists `bold/italic/caps/small_caps/strike/vanish`. There is no way to set `cs_bold`, `cs_italic`, `emboss`, `imprint`, `outline`, `shadow` via `create_style` / `modify_style`. The resolver surfaces them (per v0.2 plan) but the writer can't produce them. A user reading via `resolve_effective_formatting` and trying to round-trip through `modify_style` will hit `UnknownStylePropertyError`.
- **Suggestion:** Extend both `_RUN_LEVEL_PROPS` and `_TOGGLE_PROPS` to include the six new toggles (mapping `cs_bold` → `bCs`, etc.). Add round-trip tests mirroring `test_toggle_bold_true_round_trip` for each.

---

## Medium

### M1: `add_field(p, instruction="")` and `instruction="   "` silently emit a structurally-invalid field
- **Status:** ✅ RESOLVED (Session C) — `add_field` raises `ValueError` on empty / whitespace-only input. Parametrized regression test (`""`, `"   "`, `"\t"`, `"\n"`).
- **Subsystem:** core / fields
- **Location:** `docx_plus/fields/simple.py:129-130`
- **Description:** `wrapped = f" {instruction.strip()} "` collapses `""` or whitespace-only input to `"  "`, and `build_complex_field` happily writes `<w:instrText xml:space="preserve">  </w:instrText>`. Word renders the field as an empty result and never throws; the bug surfaces only when the user notices a blank.
- **Suggestion:** Add `if not instruction.strip(): raise ValueError("add_field requires a non-empty instruction")`. Add a regression test.

### M2: `add_page_number_field(format="")` emits a double space, not just `" PAGE "`
- **Status:** ✅ RESOLVED (Session C) — empty / whitespace-only `format` is treated as `None` (emits `" PAGE "`). Non-empty format is stripped on the way in so leading/trailing whitespace doesn't produce double spaces. Two new tests.
- **Subsystem:** core / fields
- **Location:** `docx_plus/fields/simple.py:59-62`
- **Description:** Guard is `if format is None:` not `if not format:`, so `format=""` yields `instruction = f" {field}  "` — two trailing spaces inside the field. Diverges from the documented `" PAGE "` contract.
- **Suggestion:** `instruction = f" {field} {format} " if format else f" {field} "`.

### M3: `CommentIdRegistry` does not seed from `w:commentRangeEnd`
- **Subsystem:** comments
- **Location:** `docx_plus/comments/registry.py:45-47`
- **Description:** Seeder collects ids from `w:commentRangeStart` and `w:commentReference` but skips `w:commentRangeEnd`. The docstring explicitly motivates orphan protection — a doc where rangeStart was stripped but rangeEnd remains will not block reuse of that id, defeating the stated contract.
- **Suggestion:** Add `self._collect_id_attrs(body, ".//w:commentRangeEnd")`.

### M4: `CommentNotFoundError` / `NoteNotFoundError` MRO is undocumented in `Raises` sections
- **Subsystem:** comments / notes
- **Location:** `docx_plus/comments/anchor.py:45-50`, `docx_plus/notes/write.py:42-47`
- **Description:** Dual-inheritance from `DocxPlusError` and `KeyError` works correctly (verified by tests). But public `Raises:` blocks on `edit_comment` / `edit_footnote` / `edit_endnote` don't disclose the `except KeyError` compatibility that SPEC §16 establishes.
- **Suggestion:** Add a one-line note to each `Raises` section: `"CommentNotFoundError: ... (also catches as KeyError per SPEC §16)"`.

### M5: `Border` dataclass — `color` not validated; `space` unit annotated as "twips" but is actually points
- **Subsystem:** layout
- **Location:** `docx_plus/layout/borders.py:48-78`
- **Description:** Two related defects in the same dataclass:
  - `color` accepts any string and writes verbatim. ECMA-376 17.18.79 `ST_HexColor` requires `"auto"` or six-hex-digit `"RRGGBB"`. `"red"`, `"#FF0000"`, `"FF00"`, lowercase all produce malformed XML. Other helpers validate hex (`InvalidColorError` in `styles/modify.py`).
  - Docstring says `space` is "in twips" with `24 ≈ 0.0167 inch`. ECMA-376 17.6.10 says `w:space` on `pgBorders` children is in **points** (range 0-31). Wrong unit in the documented contract.
- **Suggestion:** Add `__post_init__` validating `color` matches `^(auto|[0-9A-Fa-f]{6})$`. Correct the `space` docstring to "points (range 0-31)". Tie pairing with H7 since `space` semantics depend on `offsetFrom`.

### M6: `delete_comment` / `clear_all_comments` drops sibling content when reference-run is shared
- **Subsystem:** comments
- **Location:** `docx_plus/comments/anchor.py:236-241`
- **Description:** Cleanup finds each `<w:commentReference>` and removes its parent `<w:r>` entirely. `add_comment` builds the reference run in isolation, so internal callers are safe. But OOXML allows a `<w:r>` to contain multiple references or to mix reference with `<w:t>` text content (hand-edited / cross-tool round-tripped docs do this). Current code drops sibling text.
- **Suggestion:** Remove only the `<w:commentReference>` child; check if parent `<w:r>` is empty (or only has `<w:rPr>`) and remove conditionally. Add a test that builds a shared reference run with text content.

### M7: `set_columns` and `breaks._set_start_type` use schema-loose positioning
- **Subsystem:** layout
- **Location:** `docx_plus/layout/columns.py:94`, `docx_plus/layout/breaks.py:99-111`
- **Description:** Two pre-v0.2 helpers that violate ECMA-376 17.6.17 sectPr child order:
  - `set_columns` does `sect_pr.append(cols)`. `w:cols` belongs before `w:formProt`, `w:vAlign`, …, `w:docGrid`, `w:printerSettings`, `w:sectPrChange`. A fresh `Document()` has `w:docGrid`; the appended `w:cols` lands after it.
  - `breaks._set_start_type` does `parent.insert(0, type_el)`. `w:type` is correct only when no `w:headerReference`, `w:footerReference`, `w:footnotePr`, `w:endnotePr` precede it. Sections with custom headers/footers get `<w:type>` jammed before header refs.
- **Suggestion:** Both should use `core.insert_before_first_anchor` with correct `_LATER_SIBLINGS` tuples. Mirror the pattern from `line_numbering.py` / `borders.py`.

### M8: Outer `_apply_style_chain` doesn't read `<w:tblPr>` / `<w:trPr>` / `<w:tcPr>` from a table style's base
- **Subsystem:** styles cascade
- **Location:** `docx_plus/styles/inspect.py:455-472`
- **Description:** `_apply_style_chain` skips `<w:tblStylePr>` correctly. But the *base* of a table style may carry `<w:tblPr>`, `<w:trPr>`, `<w:tcPr>` (table/row/cell-level properties) — none of which are touched by either `_apply_style_chain` or `_apply_table_style_chain`. So cell shading (`w:shd` in `tcPr`), cell margins, row height defaults from a built-in table style are completely ignored. "Light List" declares cell shading in tcPr that the resolver won't see.
- **Suggestion:** For v0.2 release, document the limitation in `TableContext`/`resolve_effective_formatting` docstrings + CHANGELOG. For v0.3, extend the resolver to read tcPr/trPr/tblPr and add fields to `ResolvedFormatting`.

### M9: `_resolve_color` returns the theme name literally when theme is loaded but lookup fails
- **Subsystem:** styles cascade
- **Location:** `docx_plus/styles/inspect.py:823-838`
- **Description:** When `themeColor="accent7"` (unknown) is used and the theme part DID load, `resolve_theme_color` returns `None`, then `_resolve_color` sets `acc.partial = True` and returns `theme_name`. That string is stored in `color_rgb`, which downstream code (e.g. `_write_color`) rejects as invalid hex. Same logic kicks in for `themeColor="none"` which `resolve_theme_color` explicitly returns `None` for — `color_rgb = "none"` would be stored.
- **Suggestion:** In `_resolve_color`, only fall through to "return unresolved name" when the theme really IS missing (`acc.theme is None`). When theme loaded but lookup failed, return `None` (and still set `partial=True`). Special-case `name == "none"` to return `None` without setting `partial`.

### M10: `_resolve_font_theme` returns the token without setting `acc.partial = True`
- **Subsystem:** styles cascade
- **Location:** `docx_plus/styles/inspect.py:841-849`
- **Description:** Unlike `_resolve_color`, doesn't set `partial`. A document with theme font references resolves to a half-meaningful `font_name` and `partial=False`, contradicting SPEC §4. `test_doc_defaults_provide_font_name_token` even asserts the unresolved token (`"minorHAnsi"`) is returned — pinning the half-broken behaviour in tests.
- **Suggestion:** Set `acc.partial = True` whenever `_resolve_font_theme` returns a theme token; update the doc-defaults test to also assert `partial is True`. Optionally take theme part as a parameter and actually resolve `minorHAnsi`/`majorHAnsi` via `a:fontScheme`.

### M11: `find_matching_style` / `remap_styles` don't filter by `style_type` — wrong-type collisions silently break Word docs
- **Subsystem:** styles modify
- **Location:** `docx_plus/styles/modify.py:560-592` (`find_matching_style`), `:511-557` (`ensure_style`), `:667-675` (`remap_styles`)
- **Description:** A doc with a *character* style named "Heading 1" plus a `find_matching_style(doc, "Heading1")` resolves to the character style; `ensure_style(doc, "Heading1", match_existing=True)` returns that proxy; `apply_style(p, proxy.style_id)` writes `w:pStyle` pointing at a character style → Word ignores or repairs it. `remap_styles` similarly rewrites refs without checking the source style's type matches the ref type.
- **Suggestion:** Accept optional `style_type` arg on `find_matching_style`; pass `_BUILTIN_STYLES[target_id].get("style_type")` from `ensure_style`. In `remap_styles`, verify resolved style's `w:type` matches the ref tag (`pStyle` → paragraph, `rStyle` → character, `tblStyle` → table).

### M12: Body element traversal misses headers/footers/footnotes/endnotes/comments
- **Subsystem:** styles modify
- **Location:** `docx_plus/styles/modify.py:1100-1119` (`_find_references`), `:667-675` (`remap_styles`)
- **Description:** Both use `doc.part.element` as search root — main document body only. A paragraph in a header/footer referencing the style being deleted is missed; `delete_style(doc, sid)` succeeds and breaks the header on next Word open. `remap_styles` likewise leaves header/footer refs unrewritten.
- **Suggestion:** Iterate every Part that can contain WordprocessingML body content (Header, Footer, Footnotes, Endnotes, Comments) via `doc.part.related_parts`. Add coverage: insert a header, style it, attempt to delete the style → assert `StyleInUseError`.

### M13: `acc.partial = True` set upfront when theme is missing, even with no theme refs
- **Subsystem:** styles cascade
- **Location:** `docx_plus/styles/inspect.py:267-268`
- **Description:** `resolve_effective_formatting` sets `acc.partial = True` whenever the theme part is absent — but `partial` is documented as "theme resolution incomplete". On a document with no theme part AND no theme color references, the result is reported as `partial=True` even though every value is fully resolved. Makes `partial` non-actionable. No test asserts `partial is False` on an unthemed doc with no refs.
- **Suggestion:** Only set `acc.partial = True` inside `_resolve_color` / `_resolve_font_theme` when a theme reference actually fails to resolve. Drop the upfront set. Add a regression test.

### M14: `add_caption` does not apply the `Caption` paragraph style
- **Status:** ✅ RESOLVED (Session D, doc-only) — the docstring now explicitly notes the omission with a one-liner showing the caller-side recipe (`paragraph.style = doc.styles["Caption"]`). Auto-applying was rejected as too opinionated (forces a built-in dependency at every call site).
- **Subsystem:** publishing
- **Location:** `docx_plus/publishing/captions.py:67-74`, `docx_plus/examples/publishing_layout.py:39-53`
- **Description:** Word's Insert → Caption UI applies the built-in `Caption` paragraph style. This affects rendering (italic, smaller font, centered, etc., per the theme). The helper leaves the paragraph in whatever style it was created with. Captions still appear in the Table of Figures (which keys off SEQ name, not paragraph style) but render as ordinary body text — surprising for users expecting "Word-equivalent" output.
- **Suggestion:** Either auto-apply `Caption` style (perhaps gated on `apply_caption_style=True` default) or document the omission in the docstring and `ARCHITECTURE.md §7.10`. Update the example.

### M15: Example API smell — `add_caption(p, "Figure ", caption_type="Figure")` repeats label
- **Status:** ✅ RESOLVED (Session D) — `label` is now `str | None = None` with `None` defaulting to `f"{caption_type} "`. Pass `""` to suppress the label run explicitly. `docx_plus/examples/publishing_layout.py` updated to use the shorter form.
- **Subsystem:** publishing API
- **Location:** `docx_plus/examples/publishing_layout.py:40, 46, 52`; `docx_plus/publishing/captions.py:23-29`
- **Description:** Every realistic call repeats the caption-type word in both the label and the keyword. The example demonstrates the duplication three times.
- **Suggestion:** Default `label` to `f"{caption_type} "` when omitted: `def add_caption(paragraph, label: str | None = None, *, caption_type="Figure", ...)`. Keep the explicit-label path for `"Table A.", "Schedule "`. The example then collapses to `add_caption(cap1, caption_type="Figure")`.

### M16: Empty `caption_type` / `numbering` produces malformed SEQ
- **Status:** ✅ RESOLVED (Session D) — covered by H11's validation (empty `caption_type` fails the identifier regex; empty `numbering` is not a member of the picture frozenset). Both raise `ValueError` at function entry.
- **Subsystem:** publishing
- **Location:** `docx_plus/publishing/captions.py:73`, `docx_plus/publishing/figures.py:59`
- **Description:** `add_caption(p, "Figure ", caption_type="")` produces ` SEQ  \* ARABIC `. `add_caption(..., numbering="")` produces ` SEQ Figure \*  ` — a `\*` switch with no picture argument, silently dropped by Word. `numbering="lower roman"` (with space) yields `\* lower roman` → Word parses as `\* lower` with garbage trailing.
- **Suggestion:** Pair with H11 — same validation pass. Reject empty/whitespace and unknown numbering tokens at the API boundary.

### M17: `IdRangeError` is exported but undocumented in `docs/API.md`
- **Subsystem:** docs
- **Location:** `docx_plus/core/__init__.py:23,54`; missing from `docs/API.md` around line 67
- **Description:** `IdRangeError` is in `core.__all__` and SPEC §16 lists it explicitly, but `docs/API.md` lists only `DuplicateIdError` for `IdRegistry`. `docs/reference/core-ids.md` likewise omits it.
- **Suggestion:** Add a row in API.md next to `DuplicateIdError`; add it to `core-ids.md` autodoc members. (Subset of H15; resolve together.)

### M18: `mark_fields_dirty` and even/odd-header helpers only act on the first match
- **Subsystem:** core / fields / layout
- **Location:** `docx_plus/fields/update.py:62-67`, `docx_plus/layout/settings.py:84-89,101-103`
- **Description:** Both use `settings.find(qn(...))` returning only the first match. Idempotency on outputs they wrote themselves works (one element max). But settings.xml from another tool with two pathological copies leaves the second in place with stale values.
- **Suggestion:** Either document the assumption ("one element max — caller responsibility to dedupe") or change to `findall` and remove/update all matches.

### M19: `Raises` block in `insert_section_break` is incomplete
- **Subsystem:** layout
- **Location:** `docx_plus/layout/breaks.py:64-66,84-88`
- **Description:** Docstring documents one `ValueError` (paragraph not in body) but the function raises a second (`"document has no trailing sectPr to copy properties from"`) the user can't predict. Reachable on any custom-built document.
- **Suggestion:** Add to `Raises:` section: `ValueError: ... or the document has no trailing sectPr to copy section properties from.`

### M20: `notes-v0_2-expansion-scope.md` and `notes-v0_1-scope.md` referenced in shipped docs but excluded from sdist (and partially untracked)
- **Subsystem:** docs / packaging
- **Location:** `pyproject.toml:64-71` (sdist include list) vs `docx_plus/{bookmarks,comments,notes,layout,publishing}/__init__.py`, `CHANGELOG.md:16`, `docs/ARCHITECTURE.md:856`, `IMPLEMENTATION.md`
- **Description:** Eight library files link to `notes-v0_1-scope.md` and `notes-v0_2-expansion-scope.md`. `pyproject.toml` sdist `include` list omits them. PyPI users who `pip download` and read the package will see broken cross-references. Also: `git status` at review time showed `notes-v0_1-scope.md` and `notes.md` as untracked — they will not exist at a tagged commit.
- **Suggestion:** Either (a) remove the `notes-v0_*` references from library docstrings (they're internal artefacts), or (b) ship them in sdist `include`. (a) is cleaner.

### M21: `tests/conftest.py` builds fixtures into tmp dir while `build_fixtures.py main()` writes to `tests/fixtures/` — user already flagged confusion
- **Subsystem:** tests
- **Location:** `tests/conftest.py:14-18`, `tests/fixtures/build_fixtures.py:20,209,222`
- **Description:** Two divergent paths. The user's `notes.md` literally asks "What are the docx files in build/fixtures for? They seem like they may not match — do I need to manually create?" — this is the confusion.
- **Suggestion:** Pick one. If the build script is a debugging helper, make `main()` also use a tmp dir or remove the directly-writing default. If fixtures are meant to be inspectable post-build, document and remove the conftest tmp-build path. Currently neither is canonical.

### M22: `tests/test_examples_libreoffice.py` only covers 3 of 9 examples — misses the entire v0.2 surface
- **Subsystem:** tests
- **Location:** `tests/test_examples_libreoffice.py:83-90`
- **Description:** Layer-3 smoke runs against `restyle_existing`, `build_form`, `populate_form` only — leaves `add_comments`, `multi_column_layout`, `bookmarks_and_xrefs`, `footnotes_and_endnotes`, `publishing_layout` unverified by headless render. These are the v0.2 features — exactly the ones most likely to produce a "Word reports an error" file.
- **Suggestion:** Add the five v0.2 examples to the LibreOffice parametrize list. Each adds ~1 minute of CI runtime on the soffice leg.

### M23: `docs/TEST_GAPS.md` header is frozen at end-of-Phase-5 — eight IMPORTANT gaps unresolved through v0.2 with no acknowledgement
- **Subsystem:** docs
- **Location:** `docs/TEST_GAPS.md:3-5, 30-186, 224`
- **Description:** Header says "Snapshot date: 2026-05-19 (end of Phase 5)", "Suite size at snapshot: 285 tests across 17 files". Real: ~558 tests across 34 files. Recommended priority order at line 224 said "Items 1–4 are the realistic target before Phase 4 begins" — that target was missed; the file does not acknowledge the slip. Items I1–I8 remain open.
- **Suggestion:** Either close them in this release or add a top-of-file note: "v0.2 cycle did not close the IMPORTANT gaps below; they remain priority work for v0.3." Update the snapshot stats.

---

## Low

### L1: `_now_iso` strips sub-second precision
- **Subsystem:** comments
- **Location:** `docx_plus/comments/anchor.py:359-361`
- **Description:** `strftime("%Y-%m-%dT%H:%M:%SZ")` truncates microseconds. Two `add_comment` calls within the same second get identical timestamps. Not wrong (xsd:dateTime allows whole-second), but worth a fidelity decision.
- **Suggestion:** Either keep for canonical readability or switch to `.isoformat(timespec="milliseconds") + "Z"`.

### L2: `read_comments._text_between` returns nonsense on inverted ranges
- **Subsystem:** comments
- **Location:** `docx_plus/comments/read.py:152-173`
- **Description:** A malformed doc with `commentRangeEnd` before `commentRangeStart` returns `""` silently. Edge case; not covered by tests.
- **Suggestion:** No code change needed; add docstring note: "if rangeStart/rangeEnd ordering is inverted, anchored_text is empty."

### L3: `add_comment` tuple-target docstring promises validation it doesn't perform
- **Subsystem:** comments
- **Location:** `docx_plus/comments/anchor.py:91-94`, `_normalize_target:276-280`
- **Description:** "Both runs must already be parented and live in the main document body." Implementation only checks `isinstance(... Run)`; doesn't verify document order. Passing `(later_run, earlier_run)` succeeds with backwards range.
- **Suggestion:** Compare document positions and raise on inversion, or document the limitation: "caller is responsible for first appearing before second in document order."

### L4: `apply_lum_mod` / `apply_lum_off` defined but never wired into the cascade
- **Subsystem:** styles theme
- **Location:** `docx_plus/styles/theme.py:201-233`
- **Description:** Both helpers exported and unit-tested but never invoked by `_resolve_color` (which only handles `themeTint` / `themeShade`). ECMA-376 17.18.40 allows `lumMod`/`lumOff` on theme color refs; python-docx's bundled theme uses them. Any `<w:color w:themeColor="accent1" w:themeLumMod="50000"/>` resolves to unmodified accent1 with no warning.
- **Suggestion:** Wire into `_resolve_color` (small effort, fixes real-world inputs), or document as unsupported in CHANGELOG and `_resolve_color` docstring.

### L5: `_apply_cell_cascade` accepts `doc` but doesn't use it (silenced `noqa`)
- **Subsystem:** styles cascade
- **Location:** `docx_plus/styles/inspect.py:417-428`
- **Description:** `doc` param unused; `# noqa: ARG001` kept for signature symmetry with `_apply_paragraph_cascade`. Cells don't have paragraph-level numbering, so the parameter is dead.
- **Suggestion:** Drop the `doc` parameter (and the noqa). Local call sites; saves one shuffle and signals cells deliberately skip the doc-aware layers.

### L6: `target_kind` string + three `type: ignore` comments is heavier than needed
- **Subsystem:** styles cascade hygiene
- **Location:** `docx_plus/styles/inspect.py:261-292`
- **Description:** Works but is slightly fragile — a future change to `_classify_target` returning a new string value would silently take the cell branch. The variable extraction did NOT introduce aliasing problems (verified).
- **Suggestion:** Optional: have `_classify_target` return `tuple[Literal[...], etree._Element]` so mypy narrows without the ignores. Low priority.

### L7: Run target's rStyle conflated with `linkedCharStyle` in provenance
- **Subsystem:** styles cascade
- **Location:** `docx_plus/styles/inspect.py:401-414`
- **Description:** `w:link`-derived character style AND run's own `w:rStyle` reference both attributed to `"linkedCharStyle"` layer. Conceptually different; a consumer asking "why is this run pink?" can't tell which set the color.
- **Suggestion:** Introduce `"runStyle"` Layer literal; distinct from `linkedCharStyle`. Pair with C2.

### L8: `TableContext` docstring promises auto-derivation that breaks on SDT-wrapped cells
- **Subsystem:** styles cascade
- **Location:** `docx_plus/styles/inspect.py:36-69`, `_derive_table_context_from_element`
- **Description:** Works when paragraph is directly inside `<w:tc>`. If a row contains `<w:sdt>` wrapping cells, `cells.index(tc)` throws `ValueError` (caught silently, returns empty `TableContext()` — masking real positional info). Nested tables work for inner cell position only.
- **Suggestion:** Document the SDT-wrapped-cell limitation; consider walking into SDT containers when computing column index.

### L9: Self-cycle in `basedOn` raises with a slightly confusing message
- **Subsystem:** styles cascade
- **Location:** `docx_plus/styles/inspect.py:484-485`
- **Description:** Cycle detection reports the chain but can't distinguish "true cycle" from "diamond". For basedOn (single-valued) diamonds don't happen, so the message is fine — worth a code comment.
- **Suggestion:** Add a comment noting the cycle path is the actual basedOn chain; no code change.

### L10: Internal `_FootnotesPart` / `_EndnotesPart` imported by tests but not in `__all__`
- **Subsystem:** tests
- **Location:** `tests/test_core_parts.py:20-21`; `docx_plus/core/parts.py:172-178`
- **Description:** Tests reach into `parts.py` for `_FootnotesPart` / `_EndnotesPart` to assert `PartFactory.part_type_for` is wired correctly. Private (`_` prefix). If renamed or inlined, test silently breaks.
- **Suggestion:** Either expose a public predicate (`parts.registered_part_class_for(content_type)`) or add a comment in `parts.py` noting cross-reference from tests.

### L11: `xpath` recompiles every call (no compilation cache)
- **Subsystem:** core hygiene
- **Location:** `docx_plus/core/oxml.py:88`
- **Description:** `etree.XPath(expr, namespaces=NSMAP)` built fresh each call; called from hot loops (registry seeding, reads). Minor.
- **Suggestion:** Either document the choice (simplicity over caching) or add an `lru_cache` keyed on `expr`.

### L12: `test_import_invariant.py` doesn't catch relative or dynamic imports
- **Subsystem:** tests
- **Location:** `tests/test_import_invariant.py:42-51`
- **Description:** AST walk handles absolute imports only. Skips `ImportFrom` with `node.module is None` (i.e. `from . import foo`). Currently moot, but a future `from ..fields import X` from a sibling capability would pass silently. Also doesn't see `importlib.import_module`.
- **Suggestion:** Reject `level > 0` ImportFrom outright (no relative imports as a rule), or reconstruct the absolute path from `level` + file package. Document the dynamic-import gap.

### L13: `etree` import-style inconsistency across modules
- **Subsystem:** core hygiene
- **Location:** `docx_plus/fields/simple.py:20`, `bookmarks/anchor.py:21`, `comments/anchor.py:30`, `comments/read.py:19` (module-level) vs `publishing/*.py`, `core/parts.py:40` (`if TYPE_CHECKING:`)
- **Description:** Some modules import `etree` at module level despite `from __future__ import annotations`; others gate behind `TYPE_CHECKING`. Inconsistent.
- **Suggestion:** Pick one. Cost is negligible — consistency is the point.

### L14: `core/__init__.py` circular-import-by-design via `noqa: E402` pattern
- **Subsystem:** core hygiene
- **Location:** `docx_plus/core/__init__.py:23-39`
- **Description:** `noqa: E402` comments necessary because `DocxPlusError` must be defined before imports. Fine as-is but fragile — if `DocxPlusError` moves below the imports the package fails to load.
- **Suggestion:** Optional refactor — move `DocxPlusError` to its own tiny `docx_plus/core/errors.py` that no other core module imports; have `core/__init__.py` re-export.

### L15: `set_line_numbering`'s `distance` not range-validated
- **Subsystem:** layout
- **Location:** `docx_plus/layout/line_numbering.py:74-78`
- **Description:** `count_by` and `start` validated; `distance` not. Negative `distance` silently written; Word may reject.
- **Suggestion:** Add `if distance is not None and distance < 0: raise ValueError(...)` matching the existing pattern.

### L16: `read_notes` filter relies on a two-tier check that's slightly over-eager
- **Subsystem:** notes
- **Location:** `docx_plus/notes/read.py:119-126`
- **Description:** Separator entries with `id <= 0` AND/OR `w:type` filtered. A user-authored note with `w:type` set to something legal is filtered out by the type check. Probably fine for v0.2.
- **Suggestion:** No action; documenting for visibility.

### L17: `clear_all_comments` after part exists but is empty leaves comments part connected
- **Subsystem:** comments
- **Location:** `docx_plus/comments/anchor.py:182-207`
- **Description:** Docstring explicitly says "comments part is left in place (empty) so subsequent add_comment reuses it." Some Word versions complain when opening a docx with an empty comments part and the relationship intact. python-docx parser tolerates it; Word behaviour not verified across all targeted versions.
- **Suggestion:** Either verify Word 2019+ / M365 tolerate empty-part-with-relationship, or add `remove_part=True` kwarg that tears down the relationship and part itself.

### L18: `set_line_numbering` / `set_page_borders` idempotency tests don't cover pre-existing-from-load case
- **Subsystem:** tests
- **Location:** `tests/test_layout_line_numbering.py`, `tests/test_layout_borders.py`
- **Description:** Tests call the helpers twice on a fresh `Document()`. The interesting case — opening a doc that *already* has the element (from prior Word save) and calling the helper — is not covered. Replacement logic uses `sect_pr.find(...)` which works the same way, so it's a coverage gap, not a known bug.
- **Suggestion:** Add tests that pre-seed the section with the element (via `el()`) then call the setter and assert one element remains with updated attrs.

### L19: Schema-strict insertion tests use only one anchor sibling at a time
- **Subsystem:** tests
- **Location:** `tests/test_layout_line_numbering.py:99-112`, `tests/test_layout_borders.py:103-128`
- **Description:** Each test seeds one anchor (e.g. `w:cols`). The harder case — `sect_pr` with both `w:cols` and `w:docGrid` — would catch `_LATER_SIBLINGS` misordering.
- **Suggestion:** Add a test seeding `w:pgNumType`, `w:cols`, `w:docGrid` together and assert new element lands before all three in correct relative position.

### L20: Publishing helper tests don't assert the absence of `w:updateFields` (SPEC §9.1 discipline)
- **Subsystem:** tests / publishing
- **Location:** `tests/test_publishing_toc.py`, `test_publishing_captions.py`, `test_publishing_figures.py`
- **Description:** Brief specifically requires publishing helpers do NOT auto-call `mark_fields_dirty`. Tests verify the field is emitted but never verify the absence of `w:updateFields`. A regression that quietly adds `mark_fields_dirty` inside one would pass.
- **Suggestion:** Per helper, add a negative test (no flag set) and a positive test (helper + `mark_fields_dirty` → flag set).

### L21: Reference docstrings still say "deferred to v0.2" / "v0.2 goal" for items now shipped or moved
- **Subsystem:** docs / reference
- **Location:** `docs/reference/notes-write.md:7`, `docs/reference/styles-theme.md:5`, `docs/reference/protection-document.md:8-9`, `docs/reference/comments-anchor.md:11-17`
- **Description:**
  - `notes-write.md` says "in-place edits deferred to v0.3" — but `edit_footnote`/`edit_endnote` shipped in v0.2 expansion.
  - `styles-theme.md` says "Writing themes is a v0.2 goal" — v0.2 didn't ship theme writing.
  - `protection-document.md` says "Password-protected forms deferred to v0.2" — v0.2 didn't add them.
  - `comments-anchor.md` `members:` omits `edit_comment`, `clear_all_comments`, `CommentNotFoundError`.
- **Suggestion:** Audit every `docs/reference/*.md` against the current source; rewrite "deferred to v0.2" stamps; expand `members:` lists to current `__all__`.

---

## Nit

### N1: Test count number `107` (built-in styles) is fragile, repeated in 4 places
- **Subsystem:** docs / tests
- **Location:** `README.md:109-110`, `docs/ARCHITECTURE.md §5`, `IMPLEMENTATION.md §12`, `tests/test_styles_modify.py:614-624`
- **Description:** All four agree at the moment; drift risk on any future entry.
- **Suggestion:** Expose `__doc_total__` or just `len(_BUILTIN_STYLES)` from `modify.py` and reference it from one canonical location.

### N2: `add_table_of_figures` builds instruction by repeated `+=`
- **Subsystem:** publishing hygiene
- **Location:** `docx_plus/publishing/figures.py:59-63`
- **Description:** Five-line conditional concatenation followed by trailing `instruction += " "`. Easy to misread.
- **Suggestion:** Single f-string or `parts` list + `" ".join`.

### N3: Module docstring escape mismatch — `r"""` with `\\c`
- **Subsystem:** publishing hygiene
- **Location:** `docx_plus/publishing/__init__.py:15`, `captions.py:5,7,8`, `figures.py:1,4`
- **Description:** Inside `r"""..."""`, `\\` renders as `\\` (literal two backslashes), not `\`. mkdocstrings may render either way; intent is `\c`.
- **Suggestion:** Replace `\\` with `\` inside raw module docstrings, or drop the `r` prefix.

### N4: `_doc_for` duplicated between `comments/anchor.py` and `notes/write.py`
- **Subsystem:** core hygiene
- **Location:** `docx_plus/notes/write.py:281-290`, `docx_plus/comments/anchor.py:287-302`
- **Description:** Same function in both modules. SPEC §9.1 allows this in `core` since it's a shared utility.
- **Suggestion:** Move to `docx_plus/core/oxml.py` (or new `core/proxy.py`) as `body_document_for(proxy)`.

### N5: Test helper `_instruction` in `test_publishing_toc.py` is dead-defensive
- **Subsystem:** tests
- **Location:** `tests/test_publishing_toc.py:14-25`
- **Description:** Helper does `p.find(...)` then on `None` falls through to a manual loop. The fallback is what actually catches results. Compare to cleaner caption-side helper.
- **Suggestion:** Normalise to one helper in `docx_plus/_testing/ooxml_asserts.py`.

### N6: `tests/test_examples_smoke.py::assert result.stdout` is brittle
- **Subsystem:** tests
- **Location:** `tests/test_examples_smoke.py:64`
- **Description:** Requires every example to write at least one byte to stdout. Couples the smoke test to example output side-effects; an example changed to print only on `--verbose` fails.
- **Suggestion:** Drop the stdout assertion; `returncode == 0` is enough signal.

### N7: `tests/test_examples_smoke.py` EXAMPLES and WRITES_DOCX lists are manually maintained
- **Subsystem:** tests
- **Location:** `tests/test_examples_smoke.py:19-42`
- **Description:** Adding a new example requires updating both lists.
- **Suggestion:** Derive `EXAMPLES` from `pkgutil.iter_modules(docx_plus.examples.__path__)`; keep `WRITES_DOCX` manual since output-filename mapping is example-specific.

### N8: `tests/test_layout_breaks.py::test_insert_section_break_requires_body_parent` uses an unexercised `FakePart`
- **Subsystem:** tests
- **Location:** `tests/test_layout_breaks.py:123-139`
- **Description:** `FakePart()` constructed but the assertion fires before the part is consulted. Smell: the fake never runs.
- **Suggestion:** Replace with a real header-paragraph fixture if practical; otherwise add a comment explaining the fake's purpose.

### N9: `mkdocs.yml` `site_description` and `pyproject.toml` description don't mention `publishing`
- **Subsystem:** docs / packaging
- **Location:** `mkdocs.yml:2`, `pyproject.toml:4`
- **Description:** Both describe v0.1 capabilities. v0.2 added `publishing`.
- **Suggestion:** Append `, publishing` to both descriptions.

### N10: README "Build phases" table is now historic clutter
- **Subsystem:** docs
- **Location:** `README.md:344-355`
- **Description:** Contributor-facing artefact dating from v0.1. Now mostly historic; new entries added at bottom don't make it more useful.
- **Suggestion:** Collapse to "v0.1.0 / v0.2.0 — complete" or move to IMPLEMENTATION.md.

### N11: `tests/conftest.py` builds all fixtures unconditionally per session
- **Subsystem:** tests
- **Location:** `tests/conftest.py:14-18`
- **Description:** Session-scoped, ~5 small docx files built at first request. Not a correctness issue; sub-second.
- **Suggestion:** Defer to per-fixture lazy builders; or accept the overhead and move on.

### N12: `test_table_context_is_frozen` swallows any exception type
- **Subsystem:** tests
- **Location:** `tests/test_styles_table_conditional.py:89-96`
- **Description:** `except Exception: return` catches everything, not just `dataclasses.FrozenInstanceError`. Future refactor (e.g. `__slots__` → `AttributeError`) would still pass.
- **Suggestion:** Catch `dataclasses.FrozenInstanceError` specifically; mirror `test_resolved_formatting_is_frozen`.

### N13: CHANGELOG `[0.2.0]` initial-cycle entry says "Footnotes — insert-only API" without forward pointer to edit verbs
- **Subsystem:** docs / changelog
- **Location:** `CHANGELOG.md:36-41` vs `:68-71`
- **Description:** Two separated sections; a reader skimming "initial cycle" might miss that edit verbs are added below.
- **Suggestion:** Add a forward pointer or merge into one entry describing the complete shipped surface.

---

## Areas reviewed with no findings

For reference — these areas were inspected and judged clean. Future
reviewers can skip unless code in them changes:

- **Complex-field structure** (`build_complex_field`): five-run begin/instrText/separate/result/end with `xml:space="preserve"`, ECMA-376 17.16.18 compliant. All three publishing helpers route through it.
- **`mark_fields_dirty` discipline in publishing**: no publishing module imports from `docx_plus.fields`. SPEC §9.1 holds. All three docstrings tell users to call it themselves; example does.
- **Toggle property completeness on read side**: all 12 toggles wired through `_TOGGLE_RPR`, surfaced on `ResolvedFormatting`, exercised by parametric tests. Write side has H17.
- **`TableContext` non-regression**: passing `None` correctly skips conditional branch — no regression for v0.1 paragraph/run callers.
- **Reserved note id enforcement** (-1 separator, 0 continuation): correctly enforced by `_IdRegistryBase.reserve`'s 31-bit range check (`>= 1`) and `<= 0` filter in `read_notes`. Documented across `notes/registry.py`, `notes/read.py`, `notes/write.py`, `docs/API.md`. (Separator part-side entries themselves are missing — see C1.)
- **`edit_*` attribute preservation**: comment-element attributes (`author`, `date`, `initials`) and footnote/endnote-element attributes (`id`) live on the element itself, not on children. Removing child paragraphs preserves them. Verified by tests. (But see H6 for non-paragraph children.)
- **`CommentNotFoundError` / `NoteNotFoundError` KeyError compatibility**: MRO correct; tests confirm `except KeyError` catches both.
- **Error-class MRO with multiple inheritance** (`DocxPlusError, ValueError` / `…, KeyError` / `…, TypeError`): C3 linearisation works cleanly; "catch by stdlib base" path covered.
- **Import-invariant test CAPABILITIES set**: lists all 9 capabilities (`styles, controls, fields, protection, comments, layout, bookmarks, notes, publishing`).
- **All v0.2 error classes subclass `DocxPlusError`**: verified across `comments`, `notes`, `bookmarks`, `controls`, `styles`, `core`.
- **`PartFactory.part_type_for.setdefault` registration**: safe under repeated imports.
- **Examples wired into smoke tests**: all 9 example modules in `EXAMPLES` list.
- **mkdocs.yml nav vs `docs/reference/` files**: 30 reference files all listed, no extras and no missing (except SPEC.md — see C5).
- **CHANGELOG entries vs git diff `ae2abbc..HEAD`**: every bullet maps to real changes; no phantom entries.
- **`pyproject.toml` version, author, license**: consistent (`0.2.0`, `Tom Villani, PhD`, MIT). (Alpha classifier — see C5.)
- **API.md function signatures** (spot-checked `resolve_effective_formatting`, `add_caption`, `add_toc`, `set_line_numbering`, `set_page_borders`, `add_comment`, `add_bookmark`, `add_cross_reference`): match python source.
- **`add_comment` body-side marker placement** (single-run case): markers written in correct order, correct relative position.
- **`CommentIdRegistry` / `FootnoteIdRegistry` / `EndnoteIdRegistry` namespace separation**: registries properly disjoint; reserved-id checks fire correctly. (But see M3 for `commentRangeEnd` seeding gap.)
