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
