# docx_plus — v0.1 Implementation Spec

A Python library that extends python-docx with capabilities it lacks at the
OOXML level: full style cascade inspection and modification, content controls
(fillable forms), and field management. Designed to *compose* with python-docx,
not replace it: callers keep their `Document` object and use docx_plus for the
operations python-docx can't reach.

This document is the contract. Implementations that diverge from the public
API specified here are wrong, even if they work. Internal implementation is
flexible; section 9 lists the invariants that constrain it.

> **Status note (as of v0.2.0).** This is the original v0.1 design
> contract. It remains the authority on the public API shape and on the
> error taxonomy (§16, kept current through v0.2). For the *current*
> shipped surface and the live v0.3+ roadmap, see
> [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) §9 and §11 — several
> items in §15's deferred list shipped during the v0.2 cycle and are
> annotated there.

---

## 1. Purpose & Non-Goals

### What docx_plus is

The library every python-docx power user ends up writing badly: hardened
helpers for OOXML operations that sit just past python-docx's abstraction
boundary. v0.1 targets three capabilities:

- **Style cascade**: read the effective formatting that would apply to any
  paragraph/run, with provenance; modify styles in the Word-native way
  (define a style, apply it) rather than scattering direct formatting.
- **Content controls**: create text/dropdown/date/checkbox controls, read
  their values back, enforce form protection.
- **Fields**: insert simple fields (page numbers, dates), and mark fields
  dirty so Word recalculates them on next open.

### Non-goals for v0.1

The following are explicitly **out of scope** and must not be implemented
even if they seem natural extensions:

- Templating (use docxtpl)
- A `Document` subclass or any wrapper that replaces python-docx's model
- Sections, headers/footers as first-class API
- Tables beyond what python-docx already does
- Custom numbering / list definitions
- Theme manipulation (reading theme colors for resolution is in scope; writing
  themes is not)
- Comments or tracked changes API
- Password-protected forms (legacy hash algorithm; intentionally deferred)
- Data binding of controls to Custom XML Parts (deferred to v0.2)
- LibreOffice/Pages feature parity (target is Word; cross-compatibility is a
  bonus, not a requirement)

### Stylistic non-goals

- No magic attributes attached to `Document` or other python-docx objects
- No string-formatted XML; lxml only, via the helpers in `core/oxml.py`
- No silent fallbacks for malformed input — raise with a clear message

---

## 2. Architecture Overview

```
docx_plus/
├── __init__.py              # curated public API re-exports
├── core/
│   ├── __init__.py
│   ├── ns.py                # namespace constants, qn()
│   ├── oxml.py              # el(), sub(), xpath() helpers
│   ├── ids.py               # IdRegistry
│   └── parts.py             # package part / relationship helpers
├── styles/
│   ├── __init__.py
│   ├── inspect.py           # resolve_effective_formatting + ResolvedFormatting
│   ├── modify.py            # create_style, modify_style, apply_style, ...
│   └── theme.py             # theme color resolution (read-only)
├── controls/
│   ├── __init__.py
│   ├── builder.py           # FormBuilder
│   └── read.py              # read_controls, set_control_value
├── fields/
│   ├── __init__.py
│   ├── simple.py            # add_page_number_field, add_field
│   └── update.py            # mark_fields_dirty
├── protection/
│   ├── __init__.py
│   └── document.py          # protect_document, unprotect_document
├── _testing/
│   ├── __init__.py
│   └── ooxml_asserts.py     # shared assertion helpers (not public API)
├── examples/
│   ├── inspect_document.py
│   ├── restyle_existing.py
│   ├── build_form.py
│   └── populate_form.py
├── tests/
│   ├── conftest.py
│   ├── fixtures/            # small .docx files for tests
│   ├── test_styles_inspect.py
│   ├── test_styles_modify.py
│   ├── test_controls.py
│   ├── test_fields.py
│   └── test_protection.py
├── pyproject.toml
├── README.md
├── ARCHITECTURE.md
└── API.md
```

### Dependency invariant

Capability modules (`styles/`, `controls/`, `fields/`, `protection/`) depend
on `core/` only. They **must not import from each other**. If a capability
module needs functionality from another, that functionality is shared via
`core/`. This invariant is testable (a script that walks imports and
fails the build on violation) and should be enforced by a test.

### Module responsibilities

- **`core/ns.py`** — `W`, `W14`, `R`, `MC`, etc. namespace URI constants and
  the `qn(prefix:local)` shortcut that produces Clark notation.
- **`core/oxml.py`** — `el(tag)` creates a namespaced element, `sub(parent, tag)`
  creates-and-appends, `xpath(node, expr, **ns)` runs an XPath with namespace
  bindings already set. All element construction in the library goes through
  these.
- **`core/ids.py`** — `IdRegistry`, initialized from a `Document`, scans
  existing `w:id` values on construction, `.next()` issues a unique 31-bit
  integer, `.reserve(n)` reserves a specific value.
- **`core/parts.py`** — helpers for adding custom XML parts and managing
  relationships at the package level (used by controls binding in v0.2; in
  v0.1 only the relationship helpers are needed).
- **`styles/inspect.py`** — the cascade resolver.
- **`styles/modify.py`** — style creation, modification, application.
- **`styles/theme.py`** — read-only theme color resolution.
- **`controls/builder.py`** — `FormBuilder` as designed.
- **`controls/read.py`** — read and modify existing control values.
- **`fields/simple.py`** — field insertion helpers.
- **`fields/update.py`** — `mark_fields_dirty` (sets `w:updateFields` in
  settings.xml so Word recalculates on next open).
- **`protection/document.py`** — `protect_document(doc, mode=...)`.

---

## 3. Core Foundation API

### `core/ns.py`

```python
W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
R   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC  = "http://schemas.openxmlformats.org/markup-compatibility/2006"
A   = "http://schemas.openxmlformats.org/drawingml/2006/main"

NSMAP = {"w": W, "w14": W14, "r": R, "mc": MC, "a": A}

def qn(name: str) -> str:
    """'w:tag' -> '{namespace}tag' Clark notation."""
```

### `core/oxml.py`

```python
def el(tag: str, **attrs: str) -> lxml.etree._Element:
    """Create a namespaced element. Attribute keys may be namespaced
    ('w:val=...') or plain. Use this instead of OxmlElement to keep all
    construction in one place."""

def sub(parent, tag: str, **attrs: str) -> lxml.etree._Element:
    """Create a namespaced child, append to parent, return it."""

def xpath(node, expr: str) -> list:
    """Run XPath with the library's namespace map pre-bound."""

def remove(node) -> None:
    """Remove a node from its parent if it has one. No-op if detached."""
```

### `core/ids.py`

```python
class IdRegistry:
    def __init__(self, doc: Document):
        """Scan the document body and settings.xml for existing w:id values
        on sdt and similar elements. The registry tracks issued values for
        the lifetime of one document-edit session."""

    def next(self) -> int:
        """Return a new unique 31-bit positive integer id."""

    def reserve(self, value: int) -> int:
        """Reserve a specific value; raise ValueError if already issued."""

    def issued(self) -> frozenset[int]:
        """Snapshot of all issued ids."""
```

Lifecycle: created per-document, passed explicitly to functions that need it.
Do **not** attach it as a magic attribute on `Document`. Capability modules
that construct their own `IdRegistry` internally (e.g. `FormBuilder`) accept
an optional `id_registry` parameter for callers that want to share one
across multiple operations.

---

## 4. Style Inspection API — `styles/inspect.py`

The most-requested thing python-docx doesn't do, and the highest-stakes part
of v0.1. Get this right and the library lands as a power-user tool; get it
wrong and it loses credibility immediately.

### Public API

```python
def resolve_effective_formatting(
    target: Paragraph | Run | _Cell,
    *,
    include_provenance: bool = False,
) -> ResolvedFormatting:
    """Walk the OOXML formatting cascade in correct precedence order and
    return the fully-resolved properties that would apply to `target` if
    rendered now."""
```

### The cascade — implement exactly this order

Later layers win. Toggle properties XOR through the basedOn chain per
ECMA-376 17.7.3 rather than override.

1. **docDefaults** — `w:docDefaults/w:rPrDefault` and `w:pPrDefault` from
   `styles.xml`
2. **Table style** — only if `target` is inside a table; walk the table
   style's `basedOn` chain. Apply conditional formatting (`w:tblStylePr` for
   firstRow, lastRow, firstCol, etc.) based on the target's position in the
   table.
3. **Paragraph style chain** — walk the referenced paragraph style and its
   `basedOn` ancestors to the root. Cycle detection required. Max depth 11
   (Word's limit) — raise `StyleCascadeError` if exceeded.
4. **Numbering** — if the paragraph has `w:numPr`, apply formatting from the
   corresponding numbering definition in `numbering.xml`.
5. **Direct paragraph formatting** — `w:pPr` on the paragraph itself.
6. **Direct run formatting** — `w:rPr` on the run itself (runs only).

For `Run` targets, also apply the linked character style (`w:link` from the
paragraph style) before step 6.

### Toggle properties

These XOR through the cascade:

```
b, bCs, caps, emboss, i, iCs, imprint, outline, shadow, smallCaps,
strike, vanish
```

`dstrike` is **not** a toggle (per ECMA-376; verify in implementation).

Implementation note: track each toggle property's parity as the cascade is
walked. The final effective value is the parity (even count = false, odd =
true), unless an explicit `w:val="false"` is encountered, which resets the
parity to false from that point in the cascade.

### Theme references

`rFonts/@asciiTheme`, `color/@themeColor`, etc. resolve against
`word/theme/theme1.xml`. Implement read-only theme resolution in
`styles/theme.py` and call from the inspector. If the theme part is missing
or malformed, set `partial=True` on the result and return the unresolved
theme name in the relevant field.

Theme color transforms (`@themeShade`, `@themeTint`, `@lumMod`, `@lumOff`)
must be applied to the base theme color. Reference: ECMA-376 17.18.40.

### The `ResolvedFormatting` dataclass

```python
@dataclass(frozen=True)
class ResolvedFormatting:
    # Identity
    style_id: Optional[str] = None
    style_name: Optional[str] = None

    # Paragraph-level (None on Run targets unless run is whole-paragraph)
    alignment: Optional[str] = None              # 'left'|'center'|'right'|'both'|'distribute'|...
    indent_left: Optional[int] = None            # twips
    indent_right: Optional[int] = None
    indent_first_line: Optional[int] = None      # negative for hanging
    spacing_before: Optional[int] = None         # twips
    spacing_after: Optional[int] = None
    line_spacing: Optional[float] = None         # multiplier; line_spacing_rule explains
    line_spacing_rule: Optional[str] = None      # 'auto'|'exact'|'atLeast'
    keep_with_next: Optional[bool] = None
    keep_lines: Optional[bool] = None
    page_break_before: Optional[bool] = None
    outline_level: Optional[int] = None

    # Run-level (None on Paragraph targets)
    font_name: Optional[str] = None              # resolved through theme if needed
    font_size: Optional[float] = None            # points
    bold: Optional[bool] = None                  # toggle-resolved
    italic: Optional[bool] = None                # toggle-resolved
    underline: Optional[str] = None              # 'single'|'double'|'none'|...
    strike: Optional[bool] = None
    color_rgb: Optional[str] = None              # hex 'RRGGBB', resolved through theme
    highlight: Optional[str] = None
    caps: Optional[bool] = None
    small_caps: Optional[bool] = None
    vert_align: Optional[str] = None             # 'superscript'|'subscript'|None

    # Numbering
    num_id: Optional[int] = None
    num_level: Optional[int] = None

    # Meta
    partial: bool = False                        # True if theme/etc resolution incomplete
    provenance: Optional[dict[str, FormattingSource]] = None
```

### The provenance feature

When `include_provenance=True`, the `.provenance` field is populated:

```python
@dataclass(frozen=True)
class FormattingSource:
    layer: Literal[
        "docDefaults", "tableStyle", "paragraphStyle",
        "linkedCharStyle", "numbering", "directParagraph", "directRun"
    ]
    style_id: Optional[str] = None     # for *Style layers, which style provided it
    is_toggle_resolved: bool = False   # True if value came from XOR chain
    chain_depth: Optional[int] = None  # for *Style layers, how many basedOn hops
```

Every populated field in `ResolvedFormatting` gets an entry in `provenance`
keyed by the field name. Fields with `None` values are omitted from
provenance. This is the differentiated feature: it answers "why is this
paragraph 14pt italic?" and is the basis for any future debug tooling.

### Error handling

- `StyleCascadeError(Exception)` — raised on cycles in `basedOn` chains or
  depth > 11
- `MissingPartError(Exception)` — raised if a required document part is
  absent and cannot be defaulted (e.g. `numbering.xml` referenced from
  `numPr` but not present)
- Theme resolution failures are *not* exceptions; they set `partial=True`

### Test requirements

- Round-trip every cascade layer: build a doc with formatting only at that
  layer, resolve, assert it appears.
- Verify toggle XOR: 3-level style chain with bold at each level, assert
  parity.
- Verify `basedOn` cycle detection: hand-construct a cyclic chain, assert
  `StyleCascadeError`.
- Verify theme color resolution: doc with `themeColor="accent1"`, assert
  resolved hex matches `theme1.xml` accent1.
- Verify `themeShade` arithmetic: known input → known output, comparing to a
  reference value computed from the spec formula.

---

## 5. Style Modification API — `styles/modify.py`

The complement to inspection. Designed around the Word-native workflow:
define a style, apply it. Direct formatting stays python-docx's job —
docx_plus deliberately doesn't wrap it, because the user goal is "manipulate
styles," and a library that makes direct formatting equally easy fights
that goal.

### Public API

```python
def create_style(
    doc: Document,
    style_id: str,
    *,
    style_type: Literal["paragraph", "character", "table", "numbering"] = "paragraph",
    name: Optional[str] = None,                  # defaults to style_id
    based_on: Optional[str] = None,
    next_style: Optional[str] = None,
    linked_style: Optional[str] = None,
    ui_priority: int = 99,
    q_format: bool = False,                      # show in quick-style gallery
    custom: bool = True,
    **properties,                                # font_name=, font_size=, bold=, etc.
) -> StyleProxy:
    """Create a new style. Raises StyleExistsError if style_id already
    defined (use modify_style or `if_exists='replace'`)."""

def modify_style(
    doc: Document,
    style_id: str,
    *,
    if_missing: Literal["raise", "create"] = "raise",
    **properties,
) -> StyleProxy:
    """Modify a style's properties. Pass only the properties to change;
    others are preserved. Toggle properties: pass True/False to set
    explicitly (writes w:val="true"/"false"); pass None to clear the
    setting (XOR with parent will resume)."""

def apply_style(
    target: Paragraph | Run | _Cell,
    style_id: str,
) -> None:
    """Apply a style by id. Raises StyleNotFoundError if undefined."""

def ensure_style(
    doc: Document,
    style_id: str,
    *,
    match_existing: bool = False,
    **defaults_if_creating,
) -> StyleProxy:
    """Idempotent: if the style is defined (including as a latent built-in),
    return a proxy to it; if not, create it with `defaults_if_creating`.
    Never overwrites an existing definition.

    `match_existing=True` consults `find_matching_style` before falling
    back to the built-ins / custom-create path: a doc that already
    contains a style whose `w:styleId` or `w:name` matches `style_id`
    case- and space-insensitively is reused. The returned proxy's
    `style_id` may differ from the requested one; callers using
    `apply_style` should pass `proxy.style_id` (or use `remap_styles`
    for document-wide normalisation)."""

def find_matching_style(doc: Document, target_id: str) -> Optional[str]:
    """Look up an existing style whose `w:styleId` or `w:name` matches
    `target_id` case- and space-insensitively. Returns the matched
    `w:styleId`, or `None` if no defined style matches. Trivial-match
    case (exact `target_id` defined) returns `target_id`."""

def remap_styles(
    doc: Document,
    *,
    targets: Optional[list[str]] = None,
    mapping: Optional[dict[str, str]] = None,
    create_missing: bool = False,
) -> dict[str, str]:
    """Reconcile a document's styles against canonical ids. For each id
    in `targets` (defaults to the known-built-ins table), resolve by
    four-step fall-through: exact match → supplied `mapping` →
    `find_matching_style` → optional create-from-built-ins. Rewrites
    body references (`w:pStyle`, `w:rStyle`, `w:tblStyle`) in place;
    style-to-style references (`basedOn`, `next`, `link`) are left
    untouched. Returns `{target_id: resolved_id}` for every target
    resolved; unresolved targets are omitted."""

def list_styles(
    doc: Document,
    *,
    style_type: Optional[str] = None,
    include_latent: bool = False,
) -> list[StyleInfo]:
    """List defined styles. `include_latent=True` also returns built-in
    styles not yet materialized in styles.xml."""

def delete_style(doc: Document, style_id: str, *, force: bool = False) -> None:
    """Remove a style. Raises `StyleInUseError` if any paragraph/run
    references it (unless `force=True`, which leaves dangling references —
    Word will fall back to Normal)."""
```

### Properties accepted

The `**properties` kwargs accept the same field names as `ResolvedFormatting`:
`font_name`, `font_size`, `bold`, `italic`, `color_rgb`, `alignment`,
`spacing_before`, `spacing_after`, `line_spacing`, `indent_left`,
`indent_first_line`, `keep_with_next`, etc. Unrecognized keys raise
`TypeError`.

Units: `font_size` in points (float), spacing/indent in twips (int) or accept
python-docx `Length` objects. Colors as `"RRGGBB"` strings or python-docx
`RGBColor`.

### Latent styles

Word's built-in styles (Heading 1–9, TOC 1–9, List Paragraph,
PlaceholderText, etc.) are *latent*: defined by Word's defaults but not
present in `styles.xml` until used. `ensure_style` knows about them — if
asked to ensure a built-in style ID that isn't materialized, it creates the
correct definition (matching Word's defaults for that style) rather than
treating it as a new custom style.

Maintain a known-built-ins table in `styles/modify.py` covering at minimum:
`Normal`, `Heading1` through `Heading9`, `Title`, `Subtitle`, `Quote`,
`IntenseQuote`, `ListParagraph`, `Caption`, `Hyperlink`, `PlaceholderText`,
`DefaultParagraphFont`, `TableNormal`, `NoList`. Test that materialization
of each produces a style Word accepts.

### `StyleProxy`

A lightweight wrapper exposing read access to a style's properties (same
field names as `ResolvedFormatting`) and a `.element` escape hatch for
power users. Methods on it map to the same modify operations. The proxy
holds a reference to the style element, not a snapshot — reads reflect
current state.

### Errors

- `StyleExistsError`
- `StyleNotFoundError`
- `StyleInUseError`
- `StyleCascadeError` (re-used from inspect)
- `UnknownStylePropertyError` (raised for unrecognised `**properties` kwarg;
  dual-inherits `TypeError` so `except TypeError:` still catches)

### Test requirements

- Create-then-resolve: create a style with known properties, resolve a
  paragraph using it, assert the values match.
- Modify-preserves: modify one property of a multi-property style, resolve,
  assert the unchanged properties are still present.
- `ensure_style` idempotency: call twice, assert single definition in
  `styles.xml`.
- Latent materialization: `ensure_style(doc, "Heading1")` on a fresh doc
  produces a style Word recognizes (check by re-opening with python-docx
  and inspecting via the cascade resolver).
- `apply_style` round-trip: apply, resolve, get back the style's effective
  formatting.

---

## 6. Forms API — `controls/`

The implementation is essentially the `FormBuilder` already prototyped in
the docx-forms skill, ported into the library structure with light hardening.
Plus a read API.

### `controls/builder.py`

```python
class FormBuilder:
    def __init__(
        self,
        document_or_path: Optional[Document | str] = None,
        *,
        id_registry: Optional[IdRegistry] = None,
    ):
        """Wrap a Document (or open one from a path, or create a new one).
        If id_registry is None, creates one scoped to this builder."""

    doc: Document       # underlying python-docx Document

    def add_text_control(
        self, paragraph, *, tag, alias=None,
        placeholder="Click to enter text", multiline=False,
    ) -> Element: ...

    def add_dropdown(
        self, paragraph, *, tag, items, alias=None,
        placeholder="Choose an item", editable=False,
    ) -> Element: ...

    def add_date_picker(
        self, paragraph, *, tag, alias=None,
        placeholder="Click to select a date",
        date_format="M/d/yyyy", lcid="en-US",
    ) -> Element: ...

    def add_checkbox(
        self, paragraph, *, tag, alias=None, checked=False,
    ) -> Element: ...

    def save(self, path) -> str: ...
```

`items` for dropdown accepts `list[str]` or `list[tuple[str, str]]` for
display/value pairs.

The `PlaceholderText` style and unique IDs are handled internally exactly
as in the skill prototype.

### `controls/read.py`

```python
def read_controls(
    doc: Document,
    *,
    by: Literal["tag", "alias"] = "tag",
) -> dict[str, ControlValue]:
    """Return a flat dict of all content controls in the document, keyed by
    their tag (default) or alias. Repeating tags are not supported in v0.1
    (controls bound to a Custom XML Part for repeating sections are v0.2);
    if duplicate tags are encountered, raise DuplicateTagError."""

def set_control_value(
    doc: Document,
    tag: str,
    value: str | bool | datetime,
) -> None:
    """Set the value of a control by tag. Type must match control type:
    str for text/dropdown/date, bool for checkbox, datetime for date.
    Clears the placeholder state when called."""

def clear_control(doc: Document, tag: str) -> None:
    """Reset a control to its placeholder state."""

@dataclass(frozen=True)
class ControlValue:
    tag: str
    alias: Optional[str]
    control_type: Literal["text", "dropdown", "combobox", "date", "checkbox"]
    value: Optional[str | bool | datetime]
    is_placeholder: bool                # True if showing placeholder text
```

`set_control_value` for a dropdown matches against `value` first, then
`displayText`, and raises `ValueNotInListError` if neither matches (unless
the control is a combobox, in which case any string is accepted).

### Test requirements

- Build a form with every control type, save, re-open, `read_controls`
  returns every tag with `is_placeholder=True`.
- `set_control_value` then `read_controls` round-trip.
- Dropdown value-vs-display matching.
- Checkbox glyph/state sync (carry over the assertion from the skill test
  harness).
- Form-protection enforcement is correct (the schema-order check from the
  skill test harness).

---

## 7. Fields API — `fields/`

Small, focused, and worth its own module because the "Word must recalculate"
problem is workflow-shaped.

```python
# fields/simple.py
def add_page_number_field(
    paragraph,
    *,
    field: Literal["PAGE", "NUMPAGES", "SECTIONPAGES"] = "PAGE",
    format: Optional[str] = None,                # field switches, e.g. "\\* ARABIC"
) -> Element: ...

def add_date_field(
    paragraph,
    *,
    format: str = "MMMM d, yyyy",
    auto_update: bool = True,                    # DATE vs CREATEDATE semantics
) -> Element: ...

def add_field(
    paragraph,
    *,
    instruction: str,                            # raw field instruction text
    initial_text: str = "",
) -> Element:
    """Generic field insertion using the complex field syntax
    (begin/separate/end). Use this for fields without a dedicated helper."""

# fields/update.py
def mark_fields_dirty(doc: Document) -> None:
    """Set w:updateFields in settings.xml so Word recalculates all fields
    on next open. Idempotent."""
```

### Test requirements

- Insert PAGE field, save, re-open, structure intact.
- `mark_fields_dirty` idempotency: call twice, one `w:updateFields` element.
- Round-trip: open a doc with existing fields, add another, no corruption.

---

## 8. Protection API — `protection/document.py`

```python
def protect_document(
    doc: Document,
    *,
    mode: Literal["forms", "readOnly", "comments", "trackedChanges"] = "forms",
) -> None:
    """Enforce document protection. Unpassworded — provides protection
    against accidental editing, not against a determined user."""

def unprotect_document(doc: Document) -> None:
    """Remove document protection. Idempotent."""

def is_protected(doc: Document) -> bool:
    """Return True if any protection is currently enforced."""
```

Implementation must place `w:documentProtection` in the schema-correct
position within `settings.xml` (before `w:defaultTabStop`).

---

## 9. Internal Architecture Invariants

These are non-negotiable for v0.1 implementation. Tests should enforce them.

1. **No imports between capability modules.** `styles/`, `controls/`,
   `fields/`, `protection/` import from `core/` only. A test walks the AST
   and fails the build on violation.

2. **All XML element construction goes through `core/oxml.py`.** No bare
   `lxml.etree.SubElement` or `OxmlElement` calls in capability modules. No
   string-formatted XML anywhere.

3. **`IdRegistry` is the only source of new `w:id` values.** Capability
   modules that need IDs either receive a registry as a parameter or create
   one scoped to their lifetime.

4. **No magic attributes on python-docx objects.** Library state lives in
   docx_plus-owned objects (`IdRegistry`, `FormBuilder`, `StyleProxy`).

5. **All public functions have type hints.** mypy in strict mode passes on
   the library (not necessarily the tests, which can use looser hints).

6. **All public functions have docstrings.** Google-style; summary line,
   args, returns, raises. At least one example for non-trivial functions.

7. **Errors are typed for domain conditions.** Library raises subclasses
   of `DocxPlusError` (defined in `core/__init__.py`) whenever the
   condition models a meaningful domain failure — collisions, missing
   styles, malformed structure, lookup failure, cascade limits, etc.
   Raw `ValueError`/`TypeError` are permitted for **argument-shape
   validation at the public surface** (range / non-empty / wrong-type
   checks the type system already telegraphs at the boundary), which
   would only earn a typed error if catching by exception class added
   real value to callers. See §16 for the full taxonomy and the v0.2
   carve-out, and `docs/ARCHITECTURE.md` §9 for the rationale.

8. **No unrequested side effects on the input document.** Functions that
   modify state document what they modify in the docstring. `resolve_*`
   and `read_*` functions are pure reads.

---

## 10. Test Strategy

Three layers, all required.

### Layer 1: structural unit tests

Per-module test files. Each public function tested for:

- Happy path with realistic inputs
- Each documented error condition
- Edge cases listed per module above

Fast (sub-second per test), no I/O beyond reading test fixtures.

### Layer 2: round-trip tests

For every operation that modifies a document: build → save → re-open with
python-docx → assert structure survives. Catches the "Word would say
recover" class of bug that pure structural tests miss.

### Layer 3: headless render smoke tests

Gated behind `pytest.mark.requires_libreoffice` and skipped if `soffice` is
not on PATH. For each example in `examples/`, run it, convert the output to
PDF with LibreOffice headless, assert: (a) conversion succeeds with exit
code 0, (b) the PDF has the expected page count. Not a render-correctness
assertion — that needs human review — but a "does it open without errors"
smoke test.

### Shared assertion library

`_testing/ooxml_asserts.py` exports:

```python
def assert_style_defined(doc, style_id)
def assert_style_not_defined(doc, style_id)
def assert_ids_unique(doc)
def assert_no_orphan_relationships(doc)
def assert_protected(doc, mode=None)
def assert_field_dirty(doc)
def count_controls(doc, control_type=None) -> int
```

Used across test files. Internal API; not part of the public surface.

### Test fixtures

`tests/fixtures/` contains small `.docx` files for tests that need
realistic input:

- `empty.docx` — minimal valid docx
- `multistyle.docx` — defines a 3-level style chain with toggles
- `themed.docx` — uses theme colors
- `with_table.docx` — exercises table-style cascade
- `with_numbering.docx` — exercises numbering cascade
- `existing_form.docx` — pre-built form for `read_controls` tests

Fixtures generated by a `tests/fixtures/build_fixtures.py` script, not
committed pre-built — the build script is the source of truth.

### Coverage target

≥ 90% line coverage on `core/`, `styles/`, `controls/`. Lower on examples
(they're smoke-tested by Layer 3). The number is a floor not a target;
don't write degenerate tests to hit it.

---

## 11. Examples Directory

Each example is a single executable Python file with a top docstring
explaining what it demonstrates. Examples are linted, type-checked, and
smoke-tested in CI. They are *the* documentation users actually read.

### `examples/inspect_document.py`

Opens a docx, iterates paragraphs, prints each one's effective formatting
with provenance. Demonstrates `resolve_effective_formatting` with
`include_provenance=True`. Output is human-readable; format like:

```
[1] "Document Title"
    style: Title (paragraph)
    font_name: Calibri Light  <- paragraphStyle: Title (chain_depth=0)
    font_size: 28.0           <- paragraphStyle: Title
    color_rgb: 2F5496         <- paragraphStyle: Title (theme: accent1 + shade)
    bold: True                <- toggle resolved: paragraphStyle:Title XOR docDefaults
    ...
```

### `examples/restyle_existing.py`

Opens a docx, modifies the `Heading 1` style to use a different color and
size, saves. Demonstrates `modify_style` and the Word-native workflow:
changing the style changes every heading in the document.

### `examples/build_form.py`

Builds a multi-section form using `FormBuilder`. Same shape as the docx-forms
skill example.

### `examples/populate_form.py`

Opens the form built by `build_form.py`, populates every field with sample
values via `set_control_value`, saves. Demonstrates the read/write surface.

---

## 12. Documentation Requirements

Three documents, all kept in sync with the code:

- **`README.md`** — motivation, install, 60-second quickstart with one
  example each from inspection, restyle, and forms. Audience is the
  developer evaluating whether to use the library. Two pages max.

- **`API.md`** — generated from docstrings. Use `pdoc` or `mkdocs` with
  `mkdocstrings`. CI builds it on every push to main. Audience is the
  developer using the library.

- **`ARCHITECTURE.md`** — sections 2, 5 (cascade), and 9 of this spec,
  cleaned up and made present-tense. Audience is the developer extending
  or debugging the library.

Inline docstrings: Google style. Every public symbol. Non-trivial functions
include at least one example. Private helpers (leading underscore) may have
shorter docstrings.

---

## 13. Quality Gates

The library is "done" when all of these pass:

- All tests pass (`pytest`)
- `mypy --strict docx_plus/` passes
- `ruff check docx_plus/ tests/` passes (config: default rules + `D` docstring
  rules on public API, ignore on tests)
- `ruff format --check docx_plus/ tests/` passes (formatting is a separate CI
  gate from `ruff check`; the local pre-commit hook runs both — see below)
- `mkdocs build --strict` passes (enforced by the docs workflow)
- Coverage ≥ 90% on `core/`, `styles/`, `controls/`
- All four `examples/` scripts run without error
- Layer 3 smoke tests pass on a runner with LibreOffice installed
- `ARCHITECTURE.md`, `API.md`, `README.md` exist and are current
- The import-invariant test passes (no cross-capability imports)
- No `# type: ignore` without an accompanying comment explaining why

A PR that lights up CI on all of these is mergeable. A PR that doesn't is
not, regardless of how good the code looks.

**Run the lint gate locally before pushing.** `.pre-commit-config.yaml`
wires `ruff check` and `ruff format` as local hooks that shell out to
`uv run ruff`, so they use the exact ruff CI resolves — no version drift.
Install once with `uv run pre-commit install`; thereafter both run on every
commit. To check everything on demand: `uv run pre-commit run --all-files`.

---

## 14. Build & Packaging

- **Package name (PyPI)**: `docx_plus` (PyPI canonicalises to
  `docx-plus`). Verify availability before the first upload — the name
  is permanent.
- **Python**: 3.10+ (for `T | None` syntax and `Literal[...]`).
- **Runtime deps**: `python-docx>=1.0.0`, `lxml>=4.9` (transitive via
  python-docx but pin explicitly).
- **Dev deps**: `pytest`, `pytest-cov`, `mypy`, `ruff`,
  `mkdocs-material`, `mkdocstrings`, `lxml-stubs`, `pre-commit`.
- **Build system**: `hatchling` via `pyproject.toml`. No `setup.py`.
- **Typing marker**: `docx_plus/py.typed` is shipped (PEP 561) so
  downstream `mypy` users see the type hints.
- **Versioning**: SemVer. v0.1.0 is the target of this spec.
- **License**: MIT (`LICENSE` at repo root, classifier in
  `pyproject.toml`).

---

## 15. Post-v0.1 Roadmap List

The original v0.1-era list of what comes after v0.1. **This list is
historical** — it predates the v0.2 cycle and the v0.2 in-place
expansion, so it mixes items that have since shipped with items still
deferred. The authoritative current roadmap is [`ROADMAP.md`](ROADMAP.md)
at the repo root; the annotations below reconcile this list with what
actually shipped at v0.2.0.

(*Note*: `find_matching_style` and `remap_styles`, originally drafted as
v0.2 work, **landed in v0.1** as Phase 3.5. See §5 for their public API.)

- **Sections, headers, footers** as a first-class API (`sections/` module)
  — *still deferred (v0.3+).*
- **Table cell merging, borders, shading** beyond python-docx defaults
  — *still deferred (v0.3+).* (Distinct from *page* borders, which
  shipped in v0.2 as `layout/borders.py`.)
- **Custom numbering definitions** (`numbering/` module) — *still
  deferred (v0.3+).*
- **Data binding** of content controls to Custom XML Parts — *still
  deferred (v0.3+).*
- **Comments and tracked changes** read/write API — *comments **shipped
  in v0.2*** (`comments/`: add / edit / delete / clear, anchored to runs,
  paragraphs, and run ranges); **tracked changes still deferred (v0.3+).***
- **Theme manipulation** (writing themes, not just reading) — *still
  deferred (v0.3+).*
- **Glossary-based placeholder text** (the "formal" placeholder mechanism)
  — *still deferred (v0.3+).*
- **Password-protected forms** (legacy hash algorithm) — *still deferred
  (v0.3+).*
- **A high-level "restyle" planner** that takes a target `ResolvedFormatting`
  and computes the minimal cascade modification to achieve it (the inverse
  of the inspector; deferred because the design space is large) — *still
  deferred (v0.3+).*

Beyond this original list, the v0.2 cycle also shipped capabilities not
enumerated here: **footnotes / endnotes** (`notes/`), **bookmarks and
cross-references** (`bookmarks/`), **line numbering** (`layout/`), and
the **publishing primitives** (`publishing/`: TOC, captions, table of
figures). See `ARCHITECTURE.md` §11 for the consolidated picture.

---

## 16. Error Taxonomy

Every typed error in `docx_plus` subclasses `DocxPlusError` (§9.7). For
errors with a clear builtin analogue, the typed error multiple-inherits
from that builtin so existing `except ValueError:` / `except KeyError:`
clauses still catch — and so `isinstance(err, ValueError)` keeps working
for boundary-layer code.

| Error | Module | Bases | Raised by |
|---|---|---|---|
| `DocxPlusError` | `core` | `Exception` | Root; subclass-only |
| `InvalidNamespaceError` | `core.ns` | `DocxPlusError, ValueError` | `qn` on malformed name or unknown prefix |
| `DuplicateIdError` | `core.ids` | `DocxPlusError, ValueError` | `IdRegistry.reserve` on collision |
| `IdRangeError` | `core.ids` | `DocxPlusError, ValueError` | `IdRegistry.reserve` on out-of-range id |
| `StyleCascadeError` | `styles.inspect` | `DocxPlusError` | `basedOn` cycle / depth-limit overflow; or a `Run` not contained in any paragraph |
| `MissingPartError` | `styles.inspect` | `DocxPlusError` | Required part absent (reserved; no caller yet) |
| `ThemeError` | `styles.theme` | `DocxPlusError` | Structurally invalid input to theme transforms |
| `StyleExistsError` | `styles.modify` | `DocxPlusError` | `create_style` on duplicate id |
| `StyleNotFoundError` | `styles.modify` | `DocxPlusError` | Reference to undefined style id |
| `StyleInUseError` | `styles.modify` | `DocxPlusError` | `delete_style` without `force=True` on referenced style |
| `UnknownStylePropertyError` | `styles.modify` | `DocxPlusError, TypeError` | Unknown `**properties` kwarg |
| `InvalidColorError` | `styles.modify` | `DocxPlusError, ValueError` | `color_rgb` is not a valid `RRGGBB` hex string |
| `MissingNamespaceError` | `controls.builder` | `DocxPlusError` | Document root lacks `w14` |
| `InvalidDropdownItemError` | `controls.builder` | `DocxPlusError, TypeError` | `add_dropdown` item is not `str` or `(str, str)` |
| `ControlNotFoundError` | `controls.read` | `DocxPlusError, KeyError` | Tag missing in `set_control_value` / `clear_control` |
| `DuplicateTagError` | `controls.read` | `DocxPlusError, ValueError` | Two SDTs share a tag |
| `ValueNotInListError` | `controls.read` | `DocxPlusError, ValueError` | Dropdown value matches neither `w:value` nor `w:displayText` |
| `ControlTypeError` | `controls.read` | `DocxPlusError, TypeError` | `set_control_value` type mismatch |
| `CommentNotFoundError` | `comments.anchor` | `DocxPlusError, KeyError` | `edit_comment` / `delete_comment` against an id absent from `comments.xml` |
| `NoteNotFoundError` | `notes.write` | `DocxPlusError, KeyError` | `edit_footnote` / `edit_endnote` against an id absent from the relevant part |

### Raw-exception carve-out (v0.2+)

Domain failures earn a typed `DocxPlusError` subclass per §9.7.
**Argument-shape validation at the public surface** is permitted to
raise raw `ValueError` / `TypeError` directly when the condition does
not model a meaningful domain failure — i.e. the only thing a caller
could do with a typed class is restate the input invariant.

Concrete examples currently shipped:

- `IdRegistry.next()` raising `RuntimeError` on 31-bit exhaustion
  (cannot happen in practice).
- `TypeError` at the `Paragraph` / `Run` / `_Cell` boundary when a
  caller passes something else — these are programmer errors.
- `ValueError` for `set_line_numbering` / `set_columns` /
  `add_field("")` argument ranges; for invalid bookmark names; for
  "this helper only supports the main document body" guards; for
  reserved note-id checks (`<= 0`).
- `TypeError` for tuple-shape mismatches at run-range targets.

The dividing line is "does catching by class help?" — `KeyError` on a
missing comment id is a domain miss callers want to handle specifically
(→ `CommentNotFoundError`); a bookmark name that violates the W3C name
rule is an input mistake the caller would just propagate
(→ raw `ValueError`). See `docs/ARCHITECTURE.md` §9 for per-module
detail.

---

## Appendix A: References

- ECMA-376 5th Edition Part 1 (Office Open XML File Formats — Fundamentals
  and Markup Language Reference). The canonical OOXML spec.
- ECMA-376 17.7.3 — Style hierarchy and toggle properties
- ECMA-376 17.16 — Fields and hyperlinks
- ECMA-376 17.18.40 — Theme color reference
- python-docx documentation: https://python-docx.readthedocs.io/
- The docx-forms skill prototype — internal reference for the controls API
  and the test-assertion patterns

---

*End of spec.*
