# docx_plus — Implementation Notes

Companion to `SPEC.md`. The spec is the contract: what to build, what the
public API looks like, what the quality gates are. This document is the
meta-guidance: how to think about building it, what order to do it in, where
the implementation will tempt you to cut corners, and how to know when
something is actually working versus merely looking right.

Read both before starting. The spec without these notes leads to reasonable
code that drifts on the things that matter; these notes without the spec
lead to well-organized code that doesn't meet the contract.

---

## 1. Mental Model

The thing to internalize before writing any code: **OOXML correctness is the
hard part of this project, not Python design**. The algorithms are
unsurprising. The decomposition is laid out in the spec. What's hard is
producing XML that Word actually accepts and renders the way you intend —
because Word's tolerance for malformed input is asymmetric, silent failure
modes are common, and python-docx's abstractions stop right before the parts
where correctness gets subtle.

Three implications follow from this:

**Tight feedback loops matter more than usual.** Don't write 500 lines of
cascade resolver and then run the tests. Write the smallest unit, verify it
produces correct XML against a hand-checked reference, then build on it.
The cost of being wrong at the XML layer is higher than the cost of being
wrong at the API layer, because XML errors compound silently.

**Verification has to go beyond `pytest`.** Unit tests catch shape errors.
They don't catch "Word opens this file but renders it differently than you
expected" — that requires actually opening it. LibreOffice headless +
visual inspection of the rendered output is the substitute for human Word
testing on a Linux runner. Build that into the workflow from week one, not
as a polish step.

**Specs and tests precede implementation, then implementation precedes
documentation, then documentation precedes commit.** This is Command Coding
in practice: you have a spec (this doc + SPEC.md), you write the test
fixture and the test, you write the minimum implementation to pass, you
write the docstring, you commit. Inverting this order — writing
implementation first and tests after — is how subtle correctness bugs
survive.

---

## 2. Build Order

The phases below are ordered by dependency, not by importance. Don't
parallelize across phases until Phase 2 is complete; the foundation has to
be solid before the rest can rest on it.

### Phase 1: Foundation (1–2 days)

- Repository skeleton: `pyproject.toml`, directory tree from SPEC §2,
  `.gitignore`, `LICENSE`, empty `README.md`
- CI configuration: pytest, mypy strict, ruff. Get CI green on a trivial
  test before going further.
- `core/ns.py` and `core/oxml.py` with tests. These are small (~50 lines
  each) and trivial to verify. Every later module depends on them.
- `core/ids.py` with tests for: registry from empty doc, registry from doc
  with existing IDs, collision avoidance, `.reserve()` raising on
  already-issued values.
- `_testing/ooxml_asserts.py` with at least the basic asserts. Build out
  as later tests demand more.
- The import-invariant test (SPEC §9.1). Write this early — it will catch
  architectural drift the moment it happens.
- The fixture-building script in `tests/fixtures/build_fixtures.py`,
  generating at minimum `empty.docx` and `multistyle.docx`. The other
  fixtures come as tests need them.

Phase 1 is done when: `pytest` passes, `mypy --strict` passes, `ruff check`
passes, CI is green, and a placeholder integration test that builds a doc
through `core/` helpers and saves it can be opened by python-docx.

### Phase 2: Style inspection (3–5 days)

This is the largest single phase and the most consequential. Implement in
this order:

1. **`styles/theme.py`** — read-only theme color resolution. Self-contained,
   testable in isolation, needed by the cascade resolver. Implement the
   theme color lookup, then `themeShade`/`themeTint`/`lumMod`/`lumOff`
   transforms per ECMA-376 17.18.40.
2. **Cascade resolver, no provenance** — `resolve_effective_formatting`
   walking the six layers from SPEC §4. Toggle handling. Cycle detection.
   Get this completely right before adding provenance.
3. **Tests for the cascade** — at least one test per layer in isolation,
   then layered combinations, then the toggle XOR cases.
4. **Provenance pass** — modify the resolver to track provenance optionally.
   This should be a *separate* pass that doesn't change the values returned
   when `include_provenance=False`. Verify the no-provenance path still
   produces identical output.
5. **Theme resolution edge cases** — missing theme part, malformed theme,
   theme references that don't match any defined color.

Phase 2 is done when: every test in the spec's "Test requirements" for
inspect passes, the provenance output is human-readable, and
`examples/inspect_document.py` produces output that matches its docstring's
sample format on at least three test fixtures.

### Phase 3: Style modification (2–3 days)

With inspection solid, modification has a built-in verification loop:
modify a style, resolve a paragraph using it, assert the values match.
This loop catches almost all correctness bugs.

1. `create_style`, `modify_style`, `delete_style`, `apply_style`
2. `ensure_style` with the known-built-ins table. This is the trickiest
   piece — generate the built-in definitions by extracting them from a
   doc Word has materialized them in, not by guessing.
3. `StyleProxy` and `list_styles`
4. Round-trip tests for every operation

Phase 3 is done when: the create-then-resolve round-trip works for all
documented properties, latent built-in materialization works for every
style in the known-built-ins table, and `examples/restyle_existing.py`
demonstrates a meaningful restyle (e.g., change `Heading1` color across a
document with many headings).

### Phase 4: Forms (1–2 days, mostly porting)

The docx-forms skill prototype is the starting point. Port `FormBuilder`
into `controls/builder.py`, adapting it to use `IdRegistry` from `core/`.
Carry over the skill's test harness assertions into `tests/test_controls.py`.

Then write `controls/read.py` — this is new code. The `ControlValue`
dataclass and the `read_controls` / `set_control_value` / `clear_control`
trio. Test round-trips: build → save → read → modify → save → re-read.

Phase 4 is done when: every test from the skill harness passes (adapted
to the library structure), `read_controls` correctly identifies all five
control types, and the populate-then-read round-trip works for every type.

### Phase 5: Fields and protection (1 day)

Small, focused, mostly mechanical. `fields/simple.py`, `fields/update.py`,
`protection/document.py`. Idempotency tests for `mark_fields_dirty` and
`protect_document` matter — these are the obvious cases where calling
twice should not produce two elements.

### Phase 6: Polish (1–2 days)

- Complete the examples directory
- Write `README.md`, `ARCHITECTURE.md`
- Generate `API.md` (via pdoc or mkdocs-material)
- Final quality-gate sweep against SPEC §13
- Layer 3 smoke tests with LibreOffice

Total: ~10–14 working days for v0.1. Plan for 50% more — OOXML edge cases
will eat time you don't expect.

---

## 3. Day 1, Concretely

The first session is about getting the foundation right, not making
progress on features. Specifically:

1. `git init`, repo skeleton matching SPEC §2's tree (empty `__init__.py`
   files in every package, even the ones with no code yet).
2. `pyproject.toml` with dependencies, dev dependencies, build system,
   tool configs for mypy and ruff.
3. A trivial test in `tests/test_smoke.py` that imports `docx_plus` and
   asserts `True`. Run `pytest`. It should pass.
4. CI configuration (GitHub Actions, GitLab, whatever the host is). It
   should run: `pytest`, `mypy --strict docx_plus/`, `ruff check
   docx_plus/`, `ruff format --check docx_plus/`. Push, watch it go green.
5. The first real code: `core/ns.py` (~30 lines, namespace constants and
   `qn`). Its test (~20 lines, asserts the Clark notation is correct).
   Commit.
6. `core/oxml.py` (~80 lines, `el`, `sub`, `xpath`, `remove`). Its tests.
   Commit.

End of day 1: a green repository with the foundation primitives. No
features yet, but every subsequent commit builds on solid ground. This
matters because if CI breaks on day 5, you want the bisect range to be
five days, not the entire project history.

---

## 4. Implementation Patterns

A few patterns will recur across modules. Worth establishing the
conventions early and using them consistently.

### Element construction

Never construct XML elements directly with `lxml.etree.SubElement` or
python-docx's `OxmlElement`. Always go through `core/oxml.py`:

```python
from docx_plus.core.oxml import el, sub
from docx_plus.core.ns import qn

# Instead of:
# style = OxmlElement('w:style')
# style.set(qn('w:type'), 'paragraph')

# Use:
style = el("w:style", **{"w:type": "paragraph"})
```

This makes it trivial to add hooks later (validation, logging, namespace
verification) without rewriting every call site.

### Insertion order in schema-strict containers

OOXML containers like `CT_Settings` and `CT_PPr` have *required* child
ordering. Inserting an element at the wrong position produces a file Word
will silently "repair" — which sometimes works, sometimes doesn't, and is
always a bug waiting to surface.

Don't append blindly. Use the pattern from the skill's `protect()`:

```python
anchor = parent.find(qn("w:knownLaterSibling"))
if anchor is not None:
    anchor.addprevious(new_element)
else:
    parent.append(new_element)
```

Or, for elements that should come first: `parent.insert(0, new_element)`,
guarded against existing instances.

### Idempotency

Functions that modify document state should be idempotent where it makes
sense — calling `mark_fields_dirty(doc)` twice produces the same result as
calling it once. The pattern:

```python
def mark_fields_dirty(doc):
    settings = doc.settings.element
    existing = settings.find(qn("w:updateFields"))
    if existing is not None:
        existing.set(qn("w:val"), "true")
        return
    el = sub(settings, "w:updateFields", **{"w:val": "true"})
```

Test the idempotency explicitly. It's exactly the kind of thing that
breaks silently when someone "improves" the function later.

### Type hints

Use `Paragraph | Run | _Cell` style unions, not `Union[...]`. Use `Optional`
explicitly for nullable returns; don't rely on `T | None` for return types
where the user needs to think about whether `None` is meaningful. For OOXML
element types, use `lxml.etree._Element` — yes, the underscore-prefixed name
is the canonical one.

---

## 5. The Cascade Resolver, Concretely

The hardest single module in the library. A few specific recommendations:

**Implement layers in order, not all at once.** Get docDefaults working
end-to-end first (resolver + tests + a fixture that exercises only
docDefaults). Then add the paragraph style chain. Then numbering. Then
direct formatting. The order roughly tracks complexity, and each layer's
tests live in the resolver's correctness when later layers are added.

**Toggle properties are the highest-risk part.** Write a dedicated test
file for them: `tests/test_cascade_toggles.py`. Cases to cover:

- Style defines bold, no further override → effective bold
- Style A bold, style B basedOn A bold → effective *not* bold (XOR)
- Style A bold, style B basedOn A `w:b w:val="false"` → effective false
  (explicit override, not XOR)
- Direct formatting bold on paragraph using a non-bold style → effective
  bold
- Direct formatting `w:b w:val="false"` on paragraph using a bold style →
  effective not bold

These cases will look pedantic in the test file. Keep them. They are the
ones that fail in production six months later when someone tweaks the
implementation.

**Cycle detection lives in one place.** When walking `basedOn`, keep a set
of visited style IDs. If the next style is in the set, raise
`StyleCascadeError` with a clear message including the cycle path. Don't
spread cycle detection across multiple functions.

**Provenance is a *separate pass*, not interleaved.** First implement the
resolver returning values only. Then implement provenance as either (a) a
second walk that records sources, or (b) a single walk with provenance
tracking gated behind the `include_provenance` flag. Option (b) is more
efficient; option (a) is easier to verify. Pick (a) for v0.1 unless
profiling shows it matters. The cost of getting provenance wrong is a
worse debugging experience for users; the cost of getting the values wrong
is silent rendering bugs.

**Theme resolution can fail gracefully.** If `themeColor="accent1"` is
encountered but `theme1.xml` is missing or malformed, return the
unresolved theme name as the value and set `partial=True`. Don't raise.
Users running the inspector on diverse real-world documents will hit this,
and a raise turns the inspector into something they can't use.

---

## 6. Verifying OOXML Correctness

Unit tests are necessary but not sufficient. The supplementary techniques:

**Unzip and inspect.** A `.docx` is a zip. To check what your code
actually produced:

```bash
unzip -d unpacked my_test_output.docx
xmllint --format unpacked/word/document.xml | less
xmllint --format unpacked/word/styles.xml | less
```

Get fluent at this. When a test fails and you can't see why from the
assertion alone, unzip and look. The XML is the ground truth, not your
mental model of what the code should have produced.

**Diff against Word's own output.** When implementing something with
subtle correctness requirements (style materialization, complex fields,
content controls), the most reliable verification is: produce the
construct in Word, save, unzip, look at what Word wrote. Reproduce that
structure in your code. Diff your output against Word's. The places where
they differ are the places to investigate.

**LibreOffice headless conversion as a smoke test.** Running
`soffice --headless --convert-to pdf <file>` against your output catches a
broad class of "file is structurally invalid" bugs. Not a substitute for
correctness assertions, but a fast signal that the file at least parses
cleanly. Run it in CI on every example and every test fixture.

**OOXML schema validation.** The ECMA-376 schemas (`.xsd` files) can
validate your output against the strict schema. Useful for end-of-phase
verification, less useful during active development (the error messages
are dense). Set this up but don't make it part of the unit test loop —
make it part of the pre-release checklist.

---

## 7. Common Pitfalls

Things that will go wrong if you're not watching for them.

**Namespace declarations on the root.** Some elements require their
namespace to be declared on the document root, not just on the element
itself. The `w14` namespace for checkboxes is the canonical example —
python-docx's default `Document()` doesn't declare it, so adding a w14
element silently produces XML that some renderers can't parse. The skill's
`FormBuilder` handles this implicitly because python-docx 1.2.0 happens to
declare w14 — verify this is the case in the library's test suite and add
explicit declaration logic if not.

**`OxmlElement` is python-docx-private.** The `docx.oxml.OxmlElement`
helper is used throughout the python-docx codebase but is not part of its
public API. Using it directly works today and may break in the future. The
library uses `lxml.etree` directly via `core/oxml.py` to avoid this
exposure. Don't reach for `OxmlElement` even when it would be convenient.

**`paragraph._p` is python-docx-private.** Same warning. The library will
need to reach into python-docx internals occasionally — accept that, but
contain it. The pattern: capability modules access internals only through
small adapter functions in `core/`, so when python-docx changes, one place
needs updating.

**`w:id` is one of many ID-like attributes.** Don't confuse `w:id` on
SDTs with `r:id` (relationship IDs), `w:bookmarkStart/@w:id` (bookmark
IDs), `w:commentRangeStart/@w:id` (comment IDs). They're separate
namespaces with separate uniqueness requirements. `IdRegistry` in v0.1
handles only `w:id` on `w:sdt` elements. The other ID types can share the
infrastructure later but should be separate registries.

**Style ID vs style name.** The styleId (`w:styleId` attribute) is the
machine-readable identifier; the name (`w:name`) is the human-readable
display name. Word often shows the name in its UI but stores references by
ID. The library should accept IDs everywhere and treat names as a
convenience. Don't confuse them in API design — a function that takes "the
style" should take the ID, not the name.

**python-docx versions matter.** python-docx 1.0 made breaking changes
from 0.8. The library targets 1.0+. Don't rely on behavior from earlier
versions; don't assume future versions will be stable. The dependency pin
in `pyproject.toml` should be tight.

**Latent built-ins look like missing styles.** When `ensure_style(doc,
"Heading1")` is called on a fresh document, `Heading1` is not in
`styles.xml` — but it is *defined by Word's defaults*. The library has to
know about this distinction. The known-built-ins table in `styles/modify.py`
is what makes that knowledge explicit. Build it by extracting style
definitions from documents Word has materialized them in.

---

## 8. Test-Writing Guidance

Some kinds of tests pay off more than others in this library.

**High-value:** round-trip tests (build → save → read → assert), toggle
property cases, cycle detection cases, schema-order assertions (the
"`documentProtection` precedes `defaultTabStop`" style), idempotency
tests.

**Medium-value:** structural assertions (right element with right
attributes), error-condition tests (does this raise the right error?).

**Low-value:** tests that assert "calling this function doesn't crash"
without checking output. Tests that mock python-docx (don't — use real
documents).

**The shared assertion library pays for itself.** Every time you write
the same XPath query in three different tests, that's a candidate for
`_testing/ooxml_asserts.py`. Refactor early, not at the end.

**Fixtures should be generated, not committed.** The
`build_fixtures.py` script is the source of truth. Commit the script,
not the `.docx` files it produces. This means: anyone can inspect what
the fixture *is*, regenerate it if dependencies change, and verify the
fixture-building code is itself correct.

**Smoke tests for examples.** Each example script in `examples/` should
run successfully as part of the test suite. Catches the case where the
library's public API changed and the examples weren't updated.

---

## 9. When Word Rejects a File — Debugging Strategy

When LibreOffice opens a file but Word says "Word found a problem with
content," or vice versa, the workflow:

1. **Unzip the file.** `unzip -d broken broken.docx`.
2. **Validate each part.** `xmllint --noout broken/word/document.xml` —
   does it parse as XML at all? Repeat for `styles.xml`, `settings.xml`,
   `numbering.xml`.
3. **Check `[Content_Types].xml`.** Every part must be registered with
   the right content type. A new custom XML part without a content-type
   registration is silently invalid.
4. **Check `_rels/.rels` and `word/_rels/document.xml.rels`.** Every
   relationship must point to a part that exists.
5. **Compare against a known-good file.** If you have a file Word
   produced that contains the same construct, unzip it and diff. The
   difference is your bug.
6. **Schema-validate.** Run the ECMA-376 schemas against the part.
   Errors point to specific attributes or child elements.
7. **Bisect.** If you can't find it, comment out half your modifications
   and try again. Repeat until the file is accepted, then re-add the
   smallest piece that breaks it.

This procedure isn't elegant, but it's deterministic, and the bug is
always findable. The instinct to "stare at the code and hope" is the
trap. Open the file, look at the XML.

---

## 10. Definition of Done, Per Phase

Each phase has a clear exit criterion beyond "the tests pass":

- **Phase 1**: Foundation is green in CI; placeholder integration test
  builds a doc and python-docx re-opens it.
- **Phase 2**: `examples/inspect_document.py` produces useful output on
  three real-world documents (not just fixtures). The provenance feature
  works end-to-end.
- **Phase 3**: `examples/restyle_existing.py` demonstrably changes a
  document's appearance when viewed in Word/LibreOffice.
- **Phase 4**: `examples/build_form.py` and `examples/populate_form.py`
  round-trip. The form opens in Word and the controls work.
- **Phase 5**: Page numbers in a generated document show the correct
  values after Word opens it (verifies `mark_fields_dirty`).
- **Phase 6**: Every gate in SPEC §13 is green.

Don't move on from a phase until its exit criterion is met. The
temptation is to leave loose ends to fix later; the cost is debugging
multi-phase issues without knowing which phase introduced them.

---

## 11. What Success Looks Like at v0.1

A user encountering the library should be able to:

- `pip install <package>`, no surprises in dependencies
- Skim the README in 60 seconds and understand what the library does
- Run any example without modification, on any platform with Word or
  LibreOffice installed
- Use `resolve_effective_formatting` on a real-world document and get
  output that helps them understand why a paragraph looks the way it does
- Use `modify_style` to restyle a document in the Word-native way, see
  the change reflected when the document is opened in Word
- Build a fillable form with `FormBuilder`, distribute it, and have
  recipients fill it in Word without issues
- Read the API docs and find every public function documented with at
  least one example

Anything less than this is incomplete v0.1 and should not ship.

Anything beyond this is v0.2 and should not ship in v0.1.

Hold the line in both directions.

---

## 12. Progress Log

Tracks state across multi-session work. Each entry: date, phase, what was
done, what's next. Most-recent at top.

### 2026-05-19 — Phase 6: Polish — complete

- **Pre-flight scope check.** Reviewed `notes-v0_1-scope.md` (an untracked
  discussion artifact about pulling `layout/`, `comments/`, `bookmarks/`,
  `notes/` forward from v0.2). Decision: ship 0.1.0 as specced; the
  notes file's own recommendation. Phase 6 stayed narrowly scoped.
- **Survey finding.** README + `docs/ARCHITECTURE.md` + `docs/API.md` +
  `mkdocs.yml` already covered the Phase 4+5 surface (the Phase 3.6 doc
  commit reached forward). The progress-log "Next session" callout
  understated how much was already done — task #9 (docs update) and
  task #10 (API.md re-index) reduced to small touch-ups.
- **Four runnable examples** under `docx_plus/examples/` (in-package per
  ARCHITECTURE.md §1; pyproject.toml already wires ruff/coverage exclusions):
  - `inspect_document.py` — prints effective formatting + provenance for
    every paragraph in SPEC §11 output shape. No-arg form builds a small
    in-memory demo (Title / Heading1 / body) via `apply_style`.
  - `restyle_existing.py` — modifies `Heading1` (color C00000, size 20,
    bold, new spacing) and verifies via `resolve_effective_formatting`
    that the first Heading1 paragraph picks up the change.
  - `build_form.py` — onboarding form with all five control types (text /
    multiline text / dropdown with (display,value) tuples / combobox /
    date picker / checkbox), 7 controls, `protect_document(mode="forms")`.
  - `populate_form.py` — builds the form via `build_onboarding_form`,
    fills via `set_control_value`, clears one field with `clear_control`,
    re-reads to show the mixed filled/placeholder state. The three
    docx-writing examples write to `Path.cwd()` so the smoke test's
    `subprocess.run(cwd=tmp_path)` keeps the source tree clean.
- **`tests/test_examples_smoke.py`** (7 tests): every example exits 0 with
  no args; each docx-writing example produces a file python-docx can
  reopen with non-empty paragraphs.
- **`tests/test_examples_libreoffice.py`** (3 tests, gated): converts each
  example's output to PDF via `soffice --headless --convert-to pdf`,
  asserts exit-0 and non-empty PDF. Both `pytest.mark.requires_libreoffice`
  (declared in `pyproject.toml`) and a `shutil.which`-driven `skipif` are
  applied via `pytestmark` so the suite is a no-op on dev boxes lacking
  soffice. `.github/workflows/ci.yml` installs libreoffice on the
  ubuntu/3.13 leg only — the four other matrix legs autoskip.
- **Coverage gate flipped on (TEST_GAPS B1).** `fail_under = 90` added to
  `[tool.coverage.report]` + `--cov-fail-under=90` in `ci.yml`. Closing
  the previous 77% on `styles/inspect.py` required two new test files
  before the gate could pass:
  - `tests/test_cascade_numbering.py` (7 tests) — exercises cascade
    layer 4 end-to-end via a new `build_numbered` fixture that injects a
    custom `abstractNum` (indent + bold at `lvl[0]`) + `num` (id=100)
    into python-docx's already-shipped numbering part. Provenance, the
    happy path, two unparsable-attribute branches (numId garbage / ilvl
    garbage), the unknown-numId branch, missing-w:val branch, and the
    no-numPr negative-case.
  - `tests/test_cascade_run_target.py` (5 tests) — exercises Run-target
    branches (direct rPr, paragraph-style inheritance, `w:rStyle`
    linked-char chain) and `_Cell` targets (`_apply_cell_cascade` +
    table-style walk).
  - Plus `if TYPE_CHECKING:` and `@overload` added to
    `exclude_lines` in `pyproject.toml` — standard exclusions previously
    missing. Final numbers: 91.76% aggregate, styles 90.7%, core 91%,
    controls 93%.
- **API.md / mkdocs strict.** `docs/API.md` "Phase 6 stub" paragraph
  rewritten to point at the four examples; the `../docx_plus/examples`
  link replaced with prose because mkdocs --strict treats outside-docs/
  links as broken. `mkdocs.yml`'s `strict: true` is now enabled (was
  commented out with a "Phase 6" todo); `mkdocs build --strict` is
  clean. The `docs.yml` workflow already used `--strict` per c970733.
- **README.md.** Status banner switched from "early development (v0.1
  in progress)" → "v0.1 complete"; capability list "in progress —
  see roadmap below" → present-tense; roadmap row 6 marked complete.
- **SPEC §13 type-ignore audit.** The six existing `# type: ignore`
  annotations in `styles/inspect.py` (4) and `styles/modify.py` (2)
  had mypy error categories but no human-readable justification; added
  one-line explanations per the SPEC §13 last bullet ("No # type: ignore
  without an accompanying comment explaining why").
- **TEST_GAPS.md.** B1 entry moved to the `## Resolved` section with
  date and resolution notes citing the two new test files and the
  fixture. BLOCKER section now empty.
- **Wheel + sdist build green.** `uv build` produces
  `dist/docx_plus-0.1.0.tar.gz` and `docx_plus-0.1.0-py3-none-any.whl`;
  `import docx_plus` and `docx_plus.__version__ == "0.1.0"` both confirmed
  via a fresh subprocess.
- **SPEC §13 final sweep — all 9 bullets green**:
  1. ✓ All tests pass — 304 passed, 3 skipped (LibreOffice; CI ubuntu/3.13 leg installs soffice)
  2. ✓ `mypy --strict docx_plus/` — 25 source files, no issues
  3. ✓ `ruff check docx_plus/` (+ tests/) — all checks passed
  4. ✓ Coverage ≥ 90% — 91.76% aggregate, all three packages ≥90%
  5. ✓ All four examples run without error
  6. ✓ Layer 3 smoke wired (autoskip when soffice absent)
  7. ✓ ARCHITECTURE.md, API.md, README.md exist and are current
  8. ✓ Import-invariant test passes — 12 cases
  9. ✓ No `# type: ignore` lacks justification

**v0.1 is feature-complete.** Outstanding before tagging 0.1.0: (a) push
two commits to origin/main (currently 2 ahead per `git status`), (b)
answer the two questions in `notes.md` (build/fixtures setup + the v0.2
comments/refs/endnotes interface discussion — the latter is what
`notes-v0_1-scope.md` already explores). v0.2 backlog is enumerated in
SPEC §15.

**Next session — tag 0.1.0** (post-resolution of notes.md questions) or
start v0.2 work per the recommended order in `notes-v0_1-scope.md` §5:
`comments/` → `layout/` → `bookmarks/` + cross-refs → `notes/`.

### 2026-05-19 — Phase 5: Fields and protection — complete

- **Lint debt cleared** (from Phase 3.6 / Phase 4 carry-over): `uv run ruff
  check tests/ --fix` resolved 9 of 10 issues (I001 import sorts × 5, F401
  unused imports × 2, quoted-annotation cleanups × 2). The remaining `B017`
  in `tests/test_styles_inspect.py:310` was hand-narrowed from
  `pytest.raises(Exception)` to `pytest.raises(dataclasses.FrozenInstanceError)`.
  Six test files reformatted via `ruff format`. `tests/` now passes both
  `ruff check` and `ruff format --check`.
- **`core/ns.py` xml namespace**: added `XML = "http://www.w3.org/XML/1998/
  namespace"` constant and `"xml"` key in `NSMAP`, so `qn("xml:space")`
  works. Needed by `w:instrText` (and the field result `w:t`) to keep Word
  from collapsing the surrounding whitespace on the field instruction.
  Added `test_qn_xml_namespace` to `test_core_ns.py` and updated the
  `test_nsmap_keys` set expectation.
- **`fields/simple.py`** — three public functions all routing through a
  single private `_build_complex_field(p_element, instruction, initial_text)`
  helper that emits the canonical 5-run sequence
  (begin / instrText / separate / result-text / end). Each `w:instrText`
  and the result `w:t` get `xml:space="preserve"`.
  - `add_page_number_field(paragraph, *, field="PAGE", format=None)` —
    builds `" PAGE "` (or `" {field} {format} "`); seeds initial result
    `"1"` so offline viewers see something before Word recalculates.
    Accepts `PageFieldName = Literal["PAGE", "NUMPAGES", "SECTIONPAGES"]`.
  - `add_date_field(paragraph, *, format="MMMM d, yyyy", auto_update=True)`
    — emits `DATE \@ "..."` when `auto_update=True`, `CREATEDATE \@ "..."`
    otherwise. Initial text empty (Word fills on open).
  - `add_field(paragraph, *, instruction, initial_text="")` — generic
    passthrough; strips and re-wraps the instruction in single spaces so
    callers can pass it either way.
  - All three return the begin `<w:r>` element so callers can navigate or
    relocate the field. Uses `paragraph._p` (python-docx-private — called
    out in §7 as an acceptable contained internal-API touch).
- **`fields/update.py`** — single function `mark_fields_dirty(doc)`
  matching the §4 reference (idempotent, updates existing
  `val="false"` → `"true"`). Uses anchor-before-insertion against a
  12-entry tuple of CT_Settings children that come after `w:updateFields`
  (`w:hdrShapeDefaults`, `w:footnotePr`, `w:endnotePr`, `w:compat`,
  `w:docVars`, `w:rsids`, `w:mathPr`, `w:themeFontLang`,
  `w:clrSchemeMapping`, `w:shapeDefaults`, `w:decimalSymbol`,
  `w:listSeparator`). python-docx's stock `settings.xml` includes `w:compat`
  and `w:rsids`, so blank documents hit the anchor path (verified in test).
- **`protection/document.py`** — three public functions per SPEC §8:
  - `protect_document(doc, *, mode="forms")` accepts the four `ProtectionMode`
    literals (`forms`/`readOnly`/`comments`/`trackedChanges`). Emits
    `<w:documentProtection w:edit=MODE w:enforcement="1"/>`. Idempotent: a
    second call replaces the mode rather than stacking.
  - `unprotect_document(doc)` removes the element, no-op when absent.
  - `is_protected(doc)` returns presence-only (does not introspect mode).
  - Anchor-before-insertion against an 11-entry tuple of CT_Settings
    children that come after `w:documentProtection` (`w:autoFormatOverride`,
    `w:styleLockTheme`, `w:styleLockQFSet`, `w:defaultTabStop`,
    `w:autoHyphenation`, etc.). python-docx's stock settings has
    `w:defaultTabStop`, so blank documents hit the anchor path —
    schema-correctness is verified by `test_protect_document_precedes_defaultTabStop`.
  - Both fields/update.py and protection/document.py duplicate the 4-line
    `_insert_before_first_anchor` helper privately rather than share via
    `core/`. Premature abstraction risk over DRY purity; SPEC §9.1 also
    forbids capability-to-capability imports.
- **No new error classes for Phase 5.** Both modules' inputs are
  Literal-typed; mypy catches misuse statically. Runtime misuse produces
  a structurally-valid file with a semantically-wrong field/edit attribute
  that Word will surface in its UI — consistent with SPEC's "no silent
  fallbacks" but without piling on validation that duplicates the type
  system.
- **`_testing/ooxml_asserts.py`** gained the two helpers SPEC §10 promises:
  - `assert_protected(doc, mode=None)` — verifies presence + `w:enforcement="1"`;
    optionally validates `w:edit`.
  - `assert_field_dirty(doc)` — verifies `w:updateFields val="true"`.
- **`docx_plus/fields/__init__.py`** and **`docx_plus/protection/__init__.py`**
  re-export the public surface (5 symbols + 4 symbols respectively).
  Top-level `docx_plus/__init__.py` unchanged — consistent with Phase 4
  pattern of consuming via `from docx_plus.<module> import ...`.
- **Tests**: `test_fields.py` (24 tests) + `test_protection.py` (18 tests).
  Coverage: every field type's structure assertion (instruction text, run
  count, initial result, xml:space presence), generic-field whitespace
  normalisation, two-fields-in-one-paragraph composition, save→reopen
  round-trips per field type. For protection: every mode (parametrised
  over the four Literal values), schema-position invariant, idempotency,
  mode-replacement, unprotect-after-protect, round-trip, plus negative
  tests on the assertion helper. The `test_import_invariant.py` parametrised
  test picked up `fields/` and `protection/` automatically (12 cases now,
  up from 10).
- **Round-trip smoke (transient, not committed)**: a script that does
  protect + 2 PAGE fields + mark_fields_dirty + save + reopen confirms the
  produced ~36KB `.docx` survives the full chain; `is_protected()` returns
  True, `assert_protected(mode="forms")` passes, `assert_field_dirty`
  passes, and `paragraphs[0].text` reads as `"Page 1 of 1"` (the seeded
  initial results). No LibreOffice on this dev box so the soffice headless
  smoke is deferred to Phase 6's CI plumbing.
- **Quality gates green**: pytest 285/285 (was 239, +46: 24 fields + 18
  protection + 1 ns + 3 import-invariant parametrisations), mypy --strict
  (21 source files, +3), ruff check + ruff format --check on **both**
  `docx_plus/` and `tests/` (the test-file lint debt is now gone — tests/
  is part of the gate too).

**Phase 5 exit criteria status**: SPEC §7 + §8 contracts implemented;
`mark_fields_dirty` idempotent and schema-correct; `protect_document`
schema-correct (precedes `w:defaultTabStop`); the assertion helpers from
SPEC §10 are in place. The "page numbers show correct values after Word
opens" criterion from IMPLEMENTATION.md §10 is **only partially verifiable
in CI** — the structural + python-docx round-trip is the test-suite
proxy; full verification needs Microsoft Word (or LibreOffice headless,
which Phase 6 will wire up).

**Next session — Phase 6: Polish.** Per IMPLEMENTATION.md §2 / SPEC §11 + §13:
write the four examples (`inspect_document.py`, `restyle_existing.py`,
`build_form.py`, `populate_form.py`); add Layer 3 LibreOffice-headless
smoke tests gated behind `pytest.mark.requires_libreoffice`; final
quality-gate sweep against SPEC §13 (the 90% coverage threshold from
`tool.coverage.report` needs to be enforced via `--cov-fail-under` in
CI — flagged in `docs/TEST_GAPS.md` B1); update `README.md` /
`ARCHITECTURE.md` to cover the Phase 4+5 surface (forms / fields /
protection sections); `API.md` index gets four new module entries.
Budget 1-2 days.

### 2026-05-19 — Second-pass refinement from sample doc #2 — complete

- User produced `tests/fixtures/word_samples/sample-2.docx` covering the 24
  remaining entries (Header/Footer, BodyText family, MacroText,
  BalloonText, IndexHeading, TOC3–9). All 22 of the actively-sourced
  targets now have authoritative Word values.
- Confirmations (no change needed): TOC3–9 indent progression
  480/720/960/1200/1440/1680/1920 exactly matches the previous
  240-twip-step refinement. BodyText / BodyText2 (line=480 = 2.0x) /
  BodyText3 (sz=16 = 8pt) all already correct. IndexHeading already
  carries `bold: True` (Word adds majorHAnsi theme fonts on top, but
  those aren't writer-supported and are documented as a limitation).
- Refinements:
  - **Header / Footer**: added `line_spacing: 1.0` so the header/footer
    paragraph overrides Normal's 1.08 line spacing back to single, as
    Word does. Tabs (`center@4680, right@9360`) intentionally omitted —
    no tab-stop support in `_write_paragraph_property`.
  - **MacroText / MacroTextChar**: font `Courier` → `Consolas` (Word's
    actual choice for monospace styles); added `spacing_after: 0` on
    MacroText. 9-tab progression at every 480 twips omitted (same
    limitation as Header/Footer).
  - **BalloonText / BalloonTextChar**: font `Tahoma` → `Segoe UI` (Word
    365's current default for comment balloons); size 8.0 → 9.0; added
    `line_spacing: 1.0` on BalloonText.
- Tier G header comment updated to cite the sample as authoritative.
- Gates: pytest 239/239, mypy --strict clean, ruff check + format clean.

### 2026-05-19 — Latent built-in defaults refined from Word sample — complete

- Extracted authoritative style XML from a Word-saved sample
  (`tests/fixtures/word_samples/sample-1.docx`) the user produced after applying
  each latent style to a paragraph and saving in Word 365. The sample
  materialised 60 styles, of which 23 of our 47 target latent entries.
- Refined `_BUILTIN_STYLES` in `docx_plus/styles/modify.py`:
  - **TOC2..9 indent progression**: was 220-twip step (220/440/660/880/
    1100/1320/1540/1760); now 240-twip (240/480/720/960/1200/1440/1680/
    1920) to match Word's stock pattern.
  - **TOCHeading**: added `spacing_before: 240` to match Word's stock
    spacing (kept `outline_level: 9` and `q_format: True`).
  - **TOAHeading**: removed `font_size: 12.0` — Word inherits size via the
    majorHAnsi theme font, so a literal override is wrong. Kept `bold`
    and `spacing_before: 120`.
  - **Index1**: indent 200/-200 → 240/-240 (hanging indent in twips).
  - **CommentText**: dropped `spacing_after: 0` (Word doesn't set it),
    added `line_spacing: 1.0` and `font_size: 10.0`.
  - **CommentTextChar / CommentSubjectChar**: added `font_size: 10.0`.
  - **BlockText**: indent 1440/1440 → 1152/1152, dropped
    `spacing_after: 120`, added `italic: True` and `color_rgb: "156082"`
    (Word's accent1-tracked grey-blue). Border (`pBdr`) intentionally
    omitted — our property writer doesn't model paragraph borders.
- Tier E and Tier F header comments updated to cite the Word-sample as
  the source of authority (replacing the prior "ECMA-376 pattern, refine
  later" notes).
- **Bug fix** flushed out by the sample: `TableofFigures` and
  `TableofAuthorities` were keyed under `TableOfFigures` / `TableOfAuthorities`
  (capital `O`) in `_BUILTIN_STYLES`. Word writes the styleId with a
  lowercase `o`. Renamed the keys and refined defaults from the sample
  (`spacing_after: 0` on both; `indent_left: 240, indent_first_line:
  -240` on TableofAuthorities). `IntenseQuote`/`IntenseQuoteChar` also
  refined: removed `bold` (Word doesn't set it), color `2F5496` →
  `0F4761` (Word's accent1-shaded), added `spacing_before/after: 360`,
  `indent_left/right: 864`, `alignment: "center"`. `Caption` color
  `44546A` → `0E2841` (Word's text2 theme color), added
  `line_spacing: 1.0`.
- Out of scope for this pass: Header/Footer, Body Text 1-3 family, Macro
  Text, Balloon Text, Index Heading, TOC3-9 spot checks (the 240-step is
  the source-of-truth) — these didn't materialise in the sample (Word
  only materialises a latent style after it's applied to content).
  Current entries stay as ECMA-376-pattern defaults.
- Gates: pytest 239/239 green, mypy --strict clean, ruff check + format
  clean. Visual smoke check on materialised XML confirms every refined
  entry round-trips correctly via python-docx.
- Note on `ensure_style` semantics: for styles already present in
  python-docx's bundled `default.docx` (TOCHeading, ListBullet, etc.),
  `ensure_style` is a no-op — the template wins. The refined
  `_BUILTIN_STYLES` entries take effect for stripped templates,
  third-party docs missing the style, and `remap_styles(create_missing=
  True)` paths. This is the intentional idempotency contract.

### 2026-05-15 — Built-in styles table expansion — complete

- `_BUILTIN_STYLES` in `docx_plus/styles/modify.py` grew from 22 → 107
  entries (+85). Coverage is now well past SPEC §5's "at minimum" set, into
  the territory of "every style a real Word user is likely to reach for"
  per the Exhaustive scope choice.
- **Tier A — structural essentials** (6): `NoSpacing`, `Header`,
  `HeaderChar`, `Footer`, `FooterChar`, `TableGrid`.
- **Tier B — inline emphasis (character)** (7): `Strong`, `Emphasis`,
  `IntenseEmphasis`, `SubtleEmphasis`, `IntenseReference`,
  `SubtleReference`, `BookTitle`.
- **Tier C — linked Char counterparts of paragraph styles** (13):
  `Heading1Char`–`Heading9Char`, `TitleChar`, `SubtitleChar`, `QuoteChar`,
  `IntenseQuoteChar`. These are the targets of `linked_style` on the
  paragraph-style entries; without them the link refs would dangle on docs
  that need both materialised.
- **Tier D — list paragraph variants** (18): `List`/`List2`/`List3`,
  `ListBullet`/`2`/`3`/`4`/`5`, `ListNumber`/`2`/`3`/`4`/`5`,
  `ListContinue`/`2`/`3`/`4`/`5`. The Word `numPr` placeholder child is
  intentionally omitted — the table seeds shape (basedOn=Normal + indent),
  and callers wanting actual auto-numbering should attach a numbering
  definition separately.
- **Tier E — TOC / index / table-of-* navigation** (16): `TOCHeading`,
  `TOC1`–`TOC9`, `IndexHeading`, `Index1`, `TableofFigures`,
  `TableofAuthorities`, `TOAHeading`. `TOC1`–`TOC9` use Word's stock
  240-twip per-level indent progression.
- **Tier F — footnotes / endnotes / comments / balloons** (12):
  `FootnoteText`/`Char`, `FootnoteReference`, `EndnoteText`/`Char`,
  `EndnoteReference`, `CommentText`/`Char`, `CommentReference`,
  `CommentSubject`/`Char`, `BalloonText`/`Char`. Word uses the legacy
  `annotation*` names (`annotation text`, `annotation reference`,
  `annotation subject`) — we follow that convention for `w:name` while
  exposing the modern `Comment*` style ids users actually type.
- **Tier G — misc text-block** (13): `BodyText`/`Char`, `BodyText2`/`Char`,
  `BodyText3`/`Char`, `MacroText`/`Char`, `HTMLPreformatted`/`Char`,
  `PlainText`/`Char`, `NormalIndent`, `BlockText`.
- **Defaults sourcing — important per IMPLEMENTATION.md §7 "Latent
  built-ins"**: ~40 of the 85 new entries (Header/Footer/NoSpacing/Strong/
  Emphasis/list variants/BodyText/MacroText/heading & title Char styles/
  TOCHeading/TableGrid) had structural fields extracted directly from
  python-docx's bundled `default.docx` template — these are
  Word-authoritative. The remaining ~45 (TOC1–9, footnote/endnote/comment/
  balloon family, index, table-of-*, HTMLPreformatted, PlainText,
  NormalIndent, BlockText) are truly latent (in the latentStyles block but
  never materialised) and use Word's documented defaults from ECMA-376
  conventions. **Action item**: refine these against a Word-saved sample
  doc when one is available — the user offered to produce one in MS Word.
- Theme-color attributes (`themeColor`, `themeShade`, `asciiTheme`) and
  presence-only flags (`semiHidden`, `unhideWhenUsed`) are intentionally
  not emitted — they require new property kinds in `_BUILTIN_STYLES`'s
  schema. Result: visible formatting is correct but theme-tracked colors
  resolve to plain RGB. Acceptable for v0.1; can extend if needed.
- Quality gates green: pytest 239/239 (the existing
  `test_ensure_style_all_known_builtins_succeed` iterates every entry in
  `_BUILTIN_STYLES` so all 85 new entries are exercised), mypy --strict
  (18 source files), ruff check, ruff format --check on `docx_plus/`.
- Smoke check (transient script): 32 sampled new entries `ensure_style`'d,
  saved via python-docx, reopened, and present in the resulting
  `styles.xml`. No regressions in the 239-test suite.

**Next**: when a Word-saved sample doc is provided, re-extract authoritative
defaults for the ~45 truly-latent entries to replace the ECMA-376-pattern
guesses. Then proceed to Phase 5 (Fields + protection).

### 2026-05-15 — Phase 4: Forms / content controls — complete

- `controls/builder.py` adds `FormBuilder` with `add_text_control`,
  `add_dropdown` (incl. combobox via `editable=True`), `add_date_picker`,
  `add_checkbox`, plus `save(path)`. Signatures match SPEC §6.1 verbatim
  (keyword-only after the paragraph). The class self-injects the latent
  `PlaceholderText` character style on construction and verifies the
  document root declares the `w14` namespace (raises
  `MissingNamespaceError(DocxPlusError)` if not). All XML built via
  `core/oxml.el`/`sub`; all `w:id` values issued by `core/ids.IdRegistry`
  (constructor accepts an optional registry to share between builders).
  sdtPr child order matches the docx-forms skill prototype:
  `[alias?], tag, id, [showingPlcHdr?], <type-marker>` — Word's CT_SdtPr
  schema-correct sequence.
- `controls/read.py` adds `ControlValue` (frozen dataclass), `read_controls`
  (with `by="tag" | "alias"`), `set_control_value`, `clear_control`, plus
  four typed errors: `ControlNotFoundError(DocxPlusError, KeyError)`,
  `DuplicateTagError(DocxPlusError, ValueError)`,
  `ValueNotInListError(DocxPlusError, ValueError)`,
  `ControlTypeError(DocxPlusError, TypeError)`. Type detection dispatches
  on the marker child of sdtPr; rich-text SDTs (no marker) are silently
  skipped per SPEC §6 v0.1 scope. Dropdown matching tries `w:value` then
  `w:displayText`; combobox accepts freeform fallback. The auto-prepended
  empty-value placeholder list-item is filtered out during matching so it
  cannot shadow real entries. Date values round-trip through
  `w:date/@w:fullDate` (ISO 8601); the rendered text in sdtContent uses a
  best-effort human form (full Word date-format-token translation is out
  of scope for v0.1 — the canonical machine value is `@w:fullDate`).
- `controls/__init__.py` re-exports the public surface (12 symbols).
- `_testing/ooxml_asserts.py` gains `count_controls(doc, control_type=None)`
  — uses the same `_classify_sdt` helper as `read_controls` so there's one
  source of dispatch truth.
- `tests/fixtures/build_fixtures.py` gains `build_existing_form(path)`
  which constructs three SDTs (filled text, placeholder dropdown, checked
  checkbox) **by hand without FormBuilder** so read-side tests verify
  schema tolerance. Wired into `build_all` and `conftest.py`'s
  `existing_form_docx_path` fixture.
- `tests/test_controls_builder.py` (24 tests) and
  `tests/test_controls_read.py` (26 tests) cover: every control-type
  round-trip via save→reopen; multiline text; dropdown items as strings
  vs `(display, value)` tuples; combobox vs dropdown emission; checkbox
  glyph/state sync; multi-control IdRegistry continuity; existing-id
  seeding from third-party docs; sdtPr child order assertions per
  control type; placeholder-style materialisation idempotency;
  `read_controls` placeholder-state, `by="alias"` filtering, duplicate-tag
  detection; set/clear round-trips per type; dropdown by-value vs
  by-display matching; combobox freeform passthrough;
  `ControlTypeError` for type mismatches; `ControlNotFoundError` for
  unknown tags; read on the externally-built fixture; rich-text SDT
  (no marker) silently skipped.
- Quality gates green: pytest 239/239 (was 187, +52), mypy --strict
  (18 source files, +2), ruff check, ruff format --check (all on
  `docx_plus/`). `tests/test_import_invariant.py` continues to pass —
  `controls/` imports only from `docx_plus.core` (the `PlaceholderText`
  style def is duplicated inline in `builder.py` rather than imported
  from `styles.modify`, honouring SPEC §9.1).
- **Pre-existing lint debt in test files unchanged** (Phase 3.6 flag):
  ruff still reports issues on `tests/test_core_ns.py`,
  `tests/test_styles_inspect.py`, `tests/test_styles_modify.py`,
  `tests/test_styles_theme.py`, `tests/test_theme_edge_cases.py`. Out of
  scope for Phase 4. The ruff config gates `docx_plus/` only, so the
  CI workflow still passes — but a small "lint-tests cleanup" pass is
  worth doing before Phase 5.
- **Reference artifacts**: `docx-skill-files.zip` (the source skill
  prototype) and `.skill-ref/` (extracted contents) are now in
  `.gitignore`. The `FormBuilder` design closely mirrors the prototype's
  `docx_forms.py`, with three changes: (1) ID issuance routed through
  `IdRegistry` instead of an ad-hoc set, (2) all element construction
  through `core/oxml`, (3) `protect()` deliberately omitted — protection
  belongs to Phase 5/6 `protection/document.py` per SPEC §8.
- `core/parts.py` re-confirmed unnecessary for v0.1 controls (which are
  fully inline in the document; no Custom XML Part bindings until v0.2
  repeating sections). The Phase 1 plan note is now stale; will revisit
  if Phase 5 fields needs it.

**Phase 4 exit criteria status**: SPEC §6 contract implemented, all five
control types round-trip, `read_controls` discovers all of them with
correct types, `set_control_value`/`clear_control` round-trip per type,
externally-built docs read correctly, all four quality gates pass.

**Next session — Phase 5: Fields and protection.** Per IMPLEMENTATION.md
§2 / SPEC §7 + §8: implement `fields/simple.py`
(`add_page_number_field`, `add_date_field`, `add_field`) and
`fields/update.py` (`mark_fields_dirty`); then `protection/document.py`
(`protect_form`, the form-protection enforcement that locks everything
except content controls — this is what turns a doc-with-widgets into an
actual form). The schema-order check from SPEC §6 line 583 ("Form-protection
enforcement is correct") gets validated end-to-end here. Budget 1 day.
Resolve the test-file lint debt at the start of the session if quick.

### 2026-05-15 — Phase 3.6: Documentation sweep + test gap audit — complete

- `README.md` rewritten from a 9-line stub into a real quickstart for what
  Phase 1–3.5 actually ships: motivation, install, three runnable snippets
  (inspect / modify / ensure-built-in), and a phase status table. Forms,
  fields, protection listed as not-started rather than fabricated.
- `docs/ARCHITECTURE.md` authored — present-tense reference covering layout,
  the six-layer cascade walkthrough with line citations into `inspect.py`,
  schema-strict insertion patterns (`_STYLE_CHILD_ORDER` / `_PPR_CHILD_ORDER`
  / `_RPR_CHILD_ORDER` and `_ordered_insert`), the Phase 3.5 four-step
  style remap, the `_BUILTIN_STYLES` table contents and the
  "Word-2013-defaults vs python-docx-ships-Word-2007" gotcha, the §9
  invariants, the typed-error hierarchy table, and the SPEC §10
  three-layer test strategy.
- `docs/API.md` authored — hand-curated index of every public symbol grouped
  by module, with cross-links into `ARCHITECTURE.md`. Points at the
  MkDocs-generated full reference under `docs/reference/`.
- `docs/TEST_GAPS.md` authored — the audit deliverable. Severity-tiered
  (BLOCKER / IMPORTANT / NICE-TO-HAVE) with file:line citations to every
  test that does or doesn't cover the claim. Recommended fix order at the
  bottom. Resolved-entries section ready for items to move down as gaps
  close. Top blocker: SPEC §13's 90% coverage gate is defined in
  `[tool.coverage.*]` but no `fail_under` is set and CI doesn't pass
  `--cov-fail-under`. Top importants: save→reopen round-trips missing for
  3 of 5 modify operations; schema-order assertion only on `create_style`,
  not on `modify_style`'s merge path; cycle detection only tested with
  2-node cycles; `delete_style(force=True)` only smoke-checked.
- `mkdocs.yml` at repo root configures a Material-theme MkDocs site with
  mkdocstrings (Python handler, Google docstyle, source links). Nav covers
  Home / Architecture / API Index / Reference (six per-module pages under
  `docs/reference/`) / Test Gaps. `pyproject.toml` swaps `pdoc` for
  `mkdocs` + `mkdocs-material` + `mkdocstrings[python]` in both
  `[project.optional-dependencies] dev` and `[tool.uv] dev-dependencies`.
  `.gitignore` already excluded `site/` so no change needed there.
- `docs/index.md` is a slimmed mirror of the README, tailored for the
  MkDocs landing page (the README still serves GitHub).
- `docs/reference/` — six pages, one per public-surface module
  (`core-ns`, `core-oxml`, `core-ids`, `styles-inspect`, `styles-modify`,
  `styles-theme`), each a thin context paragraph plus `::: module` directive
  for mkdocstrings to auto-render.
- No source code touched. No tests added (audit deliverable is the report,
  per the approved plan). Tests still 187/187 green; `mypy --strict` (16
  files) still green.
- **Lint regression surfaced (pre-existing, not introduced here)**: on
  `uv sync --extra dev` ruff updated to 0.15.13 and now flags 7 issues in
  test files that prior phase gates claimed were green —
  `I001` × 4 (import sort) in `tests/test_core_ns.py`,
  `tests/test_styles_inspect.py`, `tests/test_styles_modify.py`,
  `tests/test_theme_edge_cases.py`; `F401` × 2 (unused imports) in
  `tests/test_styles_inspect.py` and `tests/test_styles_theme.py`;
  `B017` × 1 (`pytest.raises(Exception)`) in
  `tests/test_styles_inspect.py:310`. Plus `ruff format` wants to
  reformat 4 test files. All trivially auto-fixable
  (`uv run ruff check --fix && uv run ruff format`) but out of scope for
  this documentation-only block. Flag for the start of the next session.

**Phase 3.6 exit criteria status**: README/ARCHITECTURE/API/TEST_GAPS in
place; MkDocs configuration committed; pdoc removed from deps. Phase 6's
"`API.md`, `ARCHITECTURE.md`, `README.md` exist and are current" gate
from SPEC §13 is met for the Phase 1–3.5 surface; those docs will need
extension as Phases 4–5 land but the structure is fixed and the
mkdocstrings handler picks up new symbols automatically.

**Next session — Phase 4: Forms** (unchanged from the prior entry). Per
IMPLEMENTATION.md §2 / SPEC §6: port `FormBuilder` from the docx-forms
skill into `controls/builder.py`, adapting it to use `IdRegistry` from
`core/`. Then write `controls/read.py` (`ControlValue`, `read_controls`,
`set_control_value`, `clear_control`). Round-trip tests per control type.
Budget 1–2 days. Worth resolving `TEST_GAPS.md` B1 (coverage threshold)
and I1 (save→reopen for modify/apply/delete) alongside or just before,
so Phase 4 lands on a tighter test floor.

### 2026-05-15 — Phase 3.5: Style remapping — complete

- `find_matching_style(doc, target_id) -> str | None` — case/space-insensitive
  lookup against both `w:styleId` and `w:name` of every defined style. Returns
  the trivial match when the id is already defined exactly. Solves the
  "doc was authored with 'Heading 1' (with space) but my code references
  'Heading1'" problem.
- `remap_styles(doc, *, targets=None, mapping=None, create_missing=False)` —
  bulk reconciliation. For each target id walks four steps: exact match →
  explicit mapping → matcher → optional create-from-built-ins. Rewrites body
  references (`w:pStyle`, `w:rStyle`, `w:tblStyle`) so subsequent `apply_style`
  works without translation. Style-to-style refs inside `styles.xml`
  (`basedOn`, `next`, `link`) are intentionally left alone — keeps the remap
  a non-destructive rewrite. When `create_missing=True` falls through, the
  new style is materialised via `_materialise_builtin` so it inherits the
  doc's customised Normal automatically through `basedOn` (no special
  inheritance plumbing needed).
- `ensure_style` gains a `match_existing=False` flag. When True it consults
  `find_matching_style` before falling back to the built-ins table /
  custom-create path. The returned proxy may have a `style_id` that differs
  from the requested one — callers using `apply_style` should pass
  `proxy.style_id` (or use `remap_styles` for document-wide normalisation).
- 18 new tests (187 total): matcher coverage on id and name with case/space
  normalisation; `remap_styles` for the four resolution steps, body-ref
  rewriting, default-targets behaviour, unresolved-omission, and the
  "create_missing inherits from customised Normal" round-trip via the
  cascade; `ensure_style` with `match_existing` True/False.
- Quality gates green: `pytest` 187/187, `mypy --strict` (16 files),
  `ruff check`, `ruff format --check`.

### 2026-05-15 — Phase 3: Style modification — complete

- `styles/modify.py` — full Phase 3 surface: `create_style`, `modify_style`,
  `apply_style`, `delete_style`, `ensure_style`, `list_styles`, plus
  `StyleProxy`, `StyleInfo`, and four typed errors (`StyleExistsError`,
  `StyleNotFoundError`, `StyleInUseError`,
  `UnknownStylePropertyError(DocxPlusError, TypeError)` so the SPEC §5
  contract that unknown kwargs raise `TypeError` survives the typed-error
  invariant from §9.7).
- Property kwargs accepted by `create_style`/`modify_style` use the same
  field names as `ResolvedFormatting`, so cascade output round-trips back
  through the modifier without translation. The supported set covers every
  SPEC §4 field documented as writeable: `font_name`, `font_size`, `bold`,
  `italic`, `underline`, `strike`, `color_rgb`, `highlight`, `caps`,
  `small_caps`, `vanish`, `vert_align`, `alignment`, `indent_*`,
  `spacing_*`, `line_spacing*`, `keep_*`, `page_break_before`,
  `outline_level`. Composite elements (`w:ind`, `w:spacing`, `w:rFonts`)
  use merge semantics — multiple kwargs targeting the same XML element
  combine rather than overwrite.
- Schema-strict child orderings for `CT_Style`, `CT_PPr`, `CT_RPr` are
  enforced by `_ordered_insert`; tests `test_*_children_ordered_correctly`
  verify the schema order survives create/modify operations. This avoids
  the silent-Word-repair failure mode called out in IMPLEMENTATION.md §4.
- Toggle properties (`bold`, `italic`, `caps`, `small_caps`, `strike`,
  `vanish`) follow SPEC §5 semantics: `True` writes presence, `False`
  writes `w:val="false"`, `None` removes the element so XOR with the
  parent style resumes. Inspect.py's reader was extended in this phase to
  honour `w:val="false"` on `keepNext` / `keepLines` / `pageBreakBefore`
  so the write→read round-trip is symmetric for those flags.
- `ensure_style` is idempotent and aware of latent built-ins. The
  known-built-ins table (`_BUILTIN_STYLES`) covers every id SPEC §5
  requires "at minimum": `Normal`, `Heading1`–`Heading9`, `Title`,
  `Subtitle`, `Quote`, `IntenseQuote`, `ListParagraph`, `Caption`,
  `Hyperlink`, `PlaceholderText`, `DefaultParagraphFont`, `TableNormal`,
  `NoList`. Built-ins materialise without `w:customStyle="1"` (they are
  not user-defined) and `Normal`/`DefaultParagraphFont`/`TableNormal`/
  `NoList` carry `w:default="1"`. Word-2013-era defaults were used for
  font sizes/colors (e.g. Heading1 = 16pt, color #2F5496); python-docx
  ships its own Word-2007 versions of most of these (Heading1 = 14pt,
  color #365F91), so `ensure_style` returns the existing definition
  unchanged on a fresh `Document()` — it only consults the built-ins
  table when the id is genuinely missing.
- `apply_style` accepts `Paragraph | Run | _Cell`. Cell support iterates
  the contained paragraphs and writes `w:pStyle` to each (matches the
  "apply style to selection" semantics from Word's UI; OOXML has no
  per-cell style reference). Type validation runs before any attribute
  access so non-targets raise `TypeError` cleanly.
- `delete_style` scans the body (`w:pStyle`, `w:rStyle`, `w:tblStyle`) and
  the styles part (`w:basedOn`, `w:next`, `w:link`, `w:numStyleLink`,
  `w:styleLink`) for inbound references. Refs from the style being
  deleted to itself are excluded so a self-referential basedOn doesn't
  block its own deletion. Headers/footers are not scanned in v0.1 —
  documented limitation; revisit when a caller exercises it.
- 63 new tests in `tests/test_styles_modify.py` (169 total). Coverage
  spans: every public function's happy path; every supported property
  via cascade round-trip; schema-order assertions for style/pPr/rPr
  children; toggle True/False/None semantics; modify-preserves-others;
  modify-clears-with-None; if_missing="create" fall-through; apply_style
  to paragraph/run/cell with replace-existing semantics; delete_style
  reference detection (paragraph + basedOn) and force=True override;
  ensure_style idempotency, latent materialisation when absent, and
  "returns existing unchanged" when python-docx already shipped it;
  list_styles type filtering and include_latent; StyleProxy modify+delete
  delegation; full save→reopen round-trip via python-docx for both a
  custom style and `Heading1`.
- Quality gates green locally: `pytest` 169/169, `mypy --strict`
  (16 files), `ruff check`, `ruff format --check`.

**Phase 3 exit criteria status**: create-then-resolve works for every
documented property (cascade round-trip is the test oracle); modify
preserves untouched properties; ensure_style is idempotent and
materialises latent built-ins from the table; apply_style round-trips
through python-docx serialisation. The `examples/restyle_existing.py`
exit criterion from IMPLEMENTATION.md §10 is deferred to Phase 6 along
with the rest of the examples directory.

**Next session — Phase 4: Forms.** Per IMPLEMENTATION.md §2 / SPEC §6:
port `FormBuilder` from the docx-forms skill into `controls/builder.py`,
adapting it to use `IdRegistry` from `core/`. Then write `controls/read.py`
(`ControlValue`, `read_controls`, `set_control_value`, `clear_control`).
Round-trip tests for each control type: build → save → re-open → read →
modify → save → re-read. Budget 1–2 days (mostly a port).

### 2026-05-15 — Phase 2: Style inspection — complete

- `styles/theme.py` — read-only theme color resolution. `load_theme()`
  reads `word/theme/theme1.xml` via the document part's `theme` relationship,
  parses `a:clrScheme` into a key → hex map. `resolve_theme_color()` handles
  the WordprocessingML name aliases (text1=dk1, background1=lt1, etc. per
  ECMA-376 17.18.97) plus `themeTint` and `themeShade` modifiers. Transforms
  for `themeTint`/`themeShade`/`lumMod`/`lumOff` are exported as standalone
  functions and verified against known input/output pairs in
  `test_styles_theme.py` (22 tests).
- `styles/inspect.py` — the cascade resolver. `resolve_effective_formatting`
  walks the six SPEC §4 layers in order: docDefaults → tableStyle → paragraph
  style chain → numbering → direct pPr → direct rPr, plus linked character
  style for `Run` targets. `_Accumulator` holds in-progress state and
  provenance side-by-side; the same walk produces both the value output and
  the optional provenance dict gated on `include_provenance`. Toggle parity
  follows ECMA-376 17.7.3 — `w:val="false"`/`"0"` resets, all other values
  XOR.
- Cycle detection / depth limit (max 11) on the basedOn walk; both raise
  `StyleCascadeError`. Theme failures (missing part, malformed XML, unknown
  name) are non-fatal: `partial=True` plus the unresolved theme name in
  `color_rgb` so debugging output stays useful (SPEC §4 "Theme references").
- Conditional table formatting (`w:tblStylePr` for firstRow/lastRow/etc.) is
  deferred — the table style chain's base pPr/rPr is applied but conditional
  variants are not. Documented inline; revisit before Phase 6 if a real
  caller exercises it.
- Theme font tokens (`majorAscii`, `minorHAnsi`, …) pass through as-is in
  `font_name` for v0.1. Resolving to actual typefaces would need
  `a:fontScheme` parsing in `theme.py`; not yet a tested requirement.
- New fixture: `themed.docx` (style with `themeColor="accent1"` +
  `themeShade="80"`) exercises the theme path end-to-end.
- Tests added (49 new, 105 total): `test_styles_theme.py` (22),
  `test_styles_inspect.py` (21), `test_cascade_toggles.py` (9 — all 5 cases
  from IMPLEMENTATION.md §5 plus parity variants), `test_cascade_provenance.py`
  (10 — including the SPEC §4 invariant that provenance flag does not
  change values), `test_theme_edge_cases.py` (9 — strip-rel and corrupt-blob
  scenarios).
- Quality gates green locally: `pytest` 105/105, `mypy --strict` (15 files),
  `ruff check`, `ruff format --check`.

**Phase 2 exit criteria status**: cascade resolver works on every layer in
isolation and in combination; toggle XOR honours ECMA-376 17.7.3; cycle and
depth-limit errors are raised; theme resolution handles the missing /
malformed / unknown-name edge cases. The provenance feature is plumbed
through end-to-end. The "`examples/inspect_document.py` produces useful
output on three real-world documents" criterion from IMPLEMENTATION.md §10
is deferred to Phase 6 (Polish) along with the rest of the examples
directory — the cascade core is ready for those examples to consume.

**Next session — Phase 3: Style modification.** Per IMPLEMENTATION.md §2,
the order is: `create_style` / `modify_style` / `delete_style` /
`apply_style` → `ensure_style` with the known-built-ins table (the trickiest
piece; extract definitions from a Word-materialised doc, don't guess) →
`StyleProxy` / `list_styles` → round-trip tests for every operation. Budget
2–3 days. The Phase 2 cascade gives Phase 3 a free round-trip verifier:
modify a style, resolve a paragraph using it, assert the values match.

### 2026-05-15 — Phase 1: Foundation — complete

- `uv` environment on Python 3.13; `requires-python = ">=3.10"` in pyproject.
- `pyproject.toml` carries hatchling build, mypy strict, ruff (default + `D`
  on package, ignored on tests), pytest, coverage configs.
- Repository skeleton matches SPEC §2 flat layout: `docx_plus/` package
  sibling to `tests/`, empty `__init__.py` files in every subpackage so
  imports resolve cleanly and the import-invariant test has real targets.
- `docx_plus/core/__init__.py` defines `DocxPlusError` (SPEC §9.7 root).
- `core/ns.py` — namespace URI constants + `qn()`.
- `core/oxml.py` — `el`, `sub`, `xpath`, `remove`. `xpath` uses
  `etree.XPath(..., namespaces=NSMAP)` so it works on both raw lxml elements
  and python-docx's `BaseOxmlElement` subclasses (whose own `xpath` method
  doesn't accept a `namespaces=` kwarg — pitfall worth remembering).
- `core/ids.py` — `IdRegistry` scans `w:sdt/w:sdtPr/w:id` on both body and
  settings parts; `next()` uses `secrets.randbelow` for unpredictability
  (not strictly required but cheap insurance); `reserve()` raises
  `DuplicateIdError(DocxPlusError, ValueError)` so SPEC §3's documented
  `ValueError` and SPEC §9.7's typed-error invariant both hold.
- `core/parts.py` — Phase 1 stub. Fleshed out in Phase 4.
- `_testing/ooxml_asserts.py` seeded with `assert_ids_unique` and
  `assert_style_defined` (only the asserts that have callers this phase).
- `tests/fixtures/build_fixtures.py` produces `empty.docx` and
  `multistyle.docx` (Base → Mid → Top paragraph style chain, with bold-XOR
  toggles for the Phase 2 cascade tests to assert against).
- Tests: smoke, ns (8 cases), oxml (8 cases), ids (10 cases including
  registry seeding + non-SDT ID exclusion), import-invariant (parametrized
  over every capability dir), integration smoke (round-trip SDT through
  `core/oxml` + `core/ids`, reopen with python-docx). 32 tests, all green
  on Python 3.13.
- Quality gates green locally: `pytest` (32/32), `mypy --strict` (13 files),
  `ruff check`, `ruff format --check`.
- `.github/workflows/ci.yml` — matrix 3.10/3.11/3.12/3.13 with uv,
  pytest + mypy --strict + ruff check + ruff format --check. Not yet
  exercised — no remote.
- `LICENSE` (MIT placeholder per SPEC §14; confirm before publishing).

**Phase 1 exit criteria status**: all met. Foundation primitives stable;
import-invariant guard in place to prevent architectural drift through the
rest of the build.

**Next session — Phase 2: Style inspection.** Per IMPLEMENTATION.md §2,
the order inside Phase 2 is: `styles/theme.py` (read-only theme color
resolution) → cascade resolver without provenance → toggle XOR tests →
provenance pass → theme edge cases. Largest single phase; budget 3–5 days.

---

*End of implementation notes.*
