# Invariants, errors, and testing

## The invariants

These are the architectural commitments. Each is enforced by a test.

1. **No imports between capability modules.** `styles/`, `controls/`,
   `fields/`, `protection/` (and the v0.2 / v0.3 capabilities) may import
   from `core/` only — never from each other. Enforced by
   `tests/test_import_invariant.py`, which walks the AST of every `.py`
   file in each capability directory and asserts no import names another
   capability. The one deliberate exception is [`cli/`](cli.md): it is
   the composition layer and imports across capabilities by design, so it
   is excluded from the invariant. [`lint/`](lint.md) is the second
   composing layer and sits in the same position.

2. **All XML element construction goes through `core/oxml.py`.** No bare
   `lxml.etree.SubElement` or `OxmlElement` calls in capability modules.
   No string-formatted XML anywhere. The convention makes it possible to
   add validation/logging hooks later without rewriting every call site.
   See [schema-strict insertion](schema-order.md).

3. **Each ID namespace has its own registry.** `IdRegistry` mints SDT
   `w:id` values; `CommentIdRegistry`, `BookmarkIdRegistry`,
   `FootnoteIdRegistry`, `EndnoteIdRegistry` mint values in their own
   uniqueness domains. All five subclass the internal
   `_IdRegistryBase` in `core/ids.py` so the
   `next` / `reserve` / `issued` mechanics live in one place;
   subclasses override `_seed_from_document` to pick up the right
   existing values. Capability modules either receive a registry as a
   parameter or construct one scoped to the call. The `r:id`
   relationship namespace is python-docx's domain and is not wrapped
   by docx_plus.

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
   `ValueError`, `TypeError`, or `KeyError` for callers that still catch
   the stdlib bases. See below.

8. **No unrequested side effects on the input document.** Functions
   that mutate document state document the mutation in the docstring.
   `resolve_*` and `read_*` functions are pure reads.

---

## Error hierarchy

Every library-raised exception subclasses `DocxPlusError`. A few also
dual-inherit a stdlib base when an existing API contract (or SPEC
sentence) calls for it.

| Exception | Bases | Raised from | Meaning |
|---|---|---|---|
| `DocxPlusError` | `Exception` | `core/__init__.py` | Root of the hierarchy. Catch this to catch every library error |
| `DuplicateIdError` | `DocxPlusError`, `ValueError` | `core/ids.py` | `IdRegistry.reserve(n)` called on an already-issued value |
| `IdRangeError` | `DocxPlusError`, `ValueError` | `core/ids.py` | A reserved id falls outside the 31-bit positive range OOXML ids must occupy |
| `InvalidNamespaceError` | `DocxPlusError`, `ValueError` | `core/ns.py` | `qn()` given a malformed name or an unknown namespace prefix |
| `StyleExistsError` | `DocxPlusError` | `styles/modify.py` | `create_style` called on an ID already defined |
| `StyleNotFoundError` | `DocxPlusError` | `styles/modify.py` | `apply_style`/`modify_style`/`delete_style` referenced an undefined ID |
| `StyleInUseError` | `DocxPlusError` | `styles/modify.py` | `delete_style` (without `force=True`) on a referenced style |
| `UnknownStylePropertyError` | `DocxPlusError`, `TypeError` | `styles/modify.py` | Unrecognised `**properties` kwarg. SPEC §5 says these raise `TypeError`; dual inheritance lets both contracts hold |
| `InvalidColorError` | `DocxPlusError`, `ValueError` | `styles/modify.py` | A `color_rgb` value on `create_style`/`modify_style` that isn't a valid `RRGGBB` hex string |
| `StyleCascadeError` | `DocxPlusError` | `styles/inspect.py` | `basedOn` chain cycles or exceeds depth 11 |
| `MissingPartError` | `DocxPlusError` | `styles/inspect.py` | A referenced part is required but absent (currently unused — see [cascade layer 4](cascade.md#six-layers-low-to-high-precedence)) |
| `ThemeError` | `DocxPlusError` | `styles/theme.py` | Structurally invalid theme input to the transform functions |
| `MissingNamespaceError` | `DocxPlusError` | `controls/builder.py` | `FormBuilder` constructed against a doc whose root doesn't declare `w14` |
| `ControlNotFoundError` | `DocxPlusError`, `KeyError` | `controls/read.py` | `set_control_value`/`clear_control` referenced a tag that doesn't exist |
| `DuplicateTagError` | `DocxPlusError`, `ValueError` | `controls/read.py` | `read_controls` found two SDTs sharing a tag (v0.1 doesn't support repeating sections) |
| `ValueNotInListError` | `DocxPlusError`, `ValueError` | `controls/read.py` | `set_control_value` against a dropdown got a value that matches no item (combobox is exempt — it accepts freeform) |
| `ControlTypeError` | `DocxPlusError`, `TypeError` | `controls/read.py` | `set_control_value` got a value whose Python type doesn't match the control type (e.g. `str` to a checkbox) |
| `InvalidDropdownItemError` | `DocxPlusError`, `TypeError` | `controls/builder.py` | A dropdown/combobox `items` entry that isn't a `str` or a `(display, value)` tuple |

`fields/` and `protection/` deliberately add **no new error classes**.
Their argument types are `Literal[...]` so mypy catches misuse
statically; runtime misuse produces a structurally-valid file with a
semantically-wrong attribute that Word surfaces in its UI. The
alternative — runtime validation duplicating the type system — would
add noise without catching real bugs.

The v0.2 modules (`comments/`, `layout/`, `bookmarks/`, `notes/`,
`publishing/`) follow the same pattern. They surface only `ValueError`
and `TypeError` for argument-shape problems (bad bookmark names,
empty paragraph targets, wrong tuple shapes for run-range targets,
out-of-range `set_line_numbering` arguments) and reuse
`DuplicateIdError` / `IdRangeError` from `core/ids.py` through their
namespace-specific registries.

The v0.2 in-place expansion added two missing-lookup errors for the
new edit verbs:

| Exception | Bases | Raised from | Meaning |
|---|---|---|---|
| `CommentNotFoundError` | `DocxPlusError`, `KeyError` | `comments/anchor.py` | `edit_comment` against an id that doesn't exist in `comments.xml` (or when the comments part itself is absent) |
| `NoteNotFoundError` | `DocxPlusError`, `KeyError` | `notes/write.py` | `edit_footnote` / `edit_endnote` against an id that doesn't exist in the corresponding part |

The dual-inheritance pattern (`DuplicateIdError`, `UnknownStylePropertyError`,
the four Phase 4 `controls/read.py` errors) exists because SPEC sentences
predating §9.7's typed-error invariant documented
`ValueError` / `TypeError` / `KeyError` as the raised type. Rather than
breaking the spec contract, both bases sit on the class — `except
ValueError` and `except DocxPlusError` both catch.

---

## Testing strategy

SPEC §10 specifies three layers:

- **Layer 1 — structural unit tests.** One file per module, fast, no
  I/O beyond reading fixtures. **2055 tests** at the end of the v0.6
  cycle (2043 pass; 12 LibreOffice round-trips skip without `soffice`),
  at 96% coverage against a 90% gate.
  Of these, 631 were collected at v0.2.0: v0.1's surface (319 tests)
  plus the v0.2 cycle — `core/parts` (13), `comments/` (35),
  `layout/` (47), `bookmarks/` + cross-refs (26), `notes/` (34),
  `styles/` table conditional (13), `publishing/` (23) — plus example
  smoke tests for the new demos, plus the regression coverage added by
  the pre-publication code/docs review (cascade correctness, schema/part
  wiring, error taxonomy, publishing validation, and the six
  newly-writable run toggles). v0.3 added the balance: `revisions/`
  (mark / read / accept-reject / settings / registry) and the `cli/`
  subcommands. v0.4 added `tests/test_comments_threads.py` (reply
  anchoring and marker ordering, thread-wide resolve / reopen, nested
  reads, foreign / malformed `commentsExtended.xml` tolerance) plus the
  `comments` CLI subcommand.
- **Layer 2 — round-trip tests.** Build → save → reopen with
  `python-docx` → assert. The high-value class for OOXML
  correctness (`IMPLEMENTATION.md §8`). Phase 5 added round-trips for
  every field type plus the protect/unprotect cycle;
  [`TEST_GAPS.md`](../TEST_GAPS.md) I1 lists the remaining gaps on the
  modify side.
- **Layer 3 — headless render smoke.** Run each example, convert to
  PDF with LibreOffice headless, assert exit-0 and page count. Gated
  on the `requires_libreoffice` pytest marker.

Test fixtures live in `tests/fixtures/build_fixtures.py` (the build
script is the source of truth, not the `.docx` files it produces —
`.gitignore` excludes the generated docx files). `empty.docx`,
`multistyle.docx`, `themed.docx`, and `existing_form.docx` are built
on demand.

Shared assertions live in `docx_plus/_testing/ooxml_asserts.py`:
`assert_ids_unique`, `assert_para_ids_unique`, `assert_style_defined`,
`count_controls`, `assert_protected`, `assert_field_dirty`. The module is
internal — not re-exported from the top-level package — and is built out
lazily as later tests demand more helpers. Of the SPEC §10 helper list,
only `assert_style_not_defined` and `assert_no_orphan_relationships`
remain unwritten.

For a frozen snapshot of where the suite has real holes, see
[`TEST_GAPS.md`](../TEST_GAPS.md).
