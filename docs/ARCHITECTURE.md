# docx_plus — Architecture

Present-tense reference for how `docx_plus` is laid out and why. This
document describes what currently exists at the end of v0.2 (Phase 6
plus the v0.2 cycle: comments, layout, bookmarks / cross-references,
footnotes / endnotes). The contract that constrains it is `SPEC.md`;
the meta-guidance on how it was built and how to extend it is
`IMPLEMENTATION.md`. Read this when you need to understand the
library's shape; read those when you need to decide what to add or how.

Audience: a developer extending or debugging `docx_plus` itself, or a user
who wants more than the README before reading source.

---

## §1 Layout

```
docx_plus/
├── __init__.py              # top-level re-exports (DocxPlusError, __version__)
├── core/                    # foundation primitives — every capability depends on these
│   ├── __init__.py          # DocxPlusError (base of all typed errors) + re-exports
│   ├── ns.py                # W, W14, R, MC, A, XML namespace constants + qn()
│   ├── oxml.py              # el(), sub(), xpath(), remove(),
│   │                        # build_complex_field, insert_before_first_anchor
│   ├── ids.py               # IdRegistry, _IdRegistryBase, DuplicateIdError
│   └── parts.py             # get_or_create_part, PartSpec, COMMENTS/FOOTNOTES/ENDNOTES_SPEC
├── styles/                  # inspect, modify, theme
│   ├── __init__.py          # re-exports every public symbol from the submodules
│   ├── inspect.py           # resolve_effective_formatting + ResolvedFormatting + FormattingSource
│   ├── modify.py            # create_style, modify_style, apply_style, delete_style,
│   │                        # ensure_style, find_matching_style, remap_styles, list_styles,
│   │                        # StyleProxy, StyleInfo, _BUILTIN_STYLES table
│   └── theme.py             # ThemeColors, load_theme, resolve_theme_color,
│                            # apply_theme_tint, apply_theme_shade, apply_lum_mod, apply_lum_off
├── controls/                # content controls (SDTs)
│   ├── __init__.py          # re-exports the public surface
│   ├── builder.py           # FormBuilder, MissingNamespaceError, DropdownItem
│   └── read.py              # ControlValue, read_controls, set_control_value, clear_control,
│                            # ControlNotFoundError, DuplicateTagError, ValueNotInListError,
│                            # ControlTypeError
├── fields/                  # complex field insertion + update flag
│   ├── __init__.py          # re-exports the public surface
│   ├── simple.py            # add_page_number_field, add_date_field, add_field,
│   │                        # PageFieldName Literal
│   └── update.py            # mark_fields_dirty
├── protection/              # document-level protection enforcement
│   ├── __init__.py          # re-exports the public surface
│   └── document.py          # protect_document, unprotect_document, is_protected,
│                            # ProtectionMode Literal
├── comments/                # anchored comments — v0.2
│   ├── __init__.py          # re-exports the public surface
│   ├── anchor.py            # add_comment, edit_comment, delete_comment, clear_all_comments,
│   │                        # CommentRef, CommentTarget, CommentNotFoundError
│   ├── read.py              # read_comments, AnchoredComment
│   └── registry.py          # CommentIdRegistry
├── layout/                  # page-layout extras — v0.2
│   ├── __init__.py          # re-exports the public surface
│   ├── columns.py           # set_columns
│   ├── breaks.py            # insert_section_break, SectionStartType
│   ├── settings.py          # enable/disable_distinct_even_odd_headers
│   ├── line_numbering.py    # set_line_numbering, LineNumberRestart
│   └── borders.py           # set_page_borders, Border
├── bookmarks/               # bookmarks + REF/PAGEREF cross-references — v0.2
│   ├── __init__.py          # re-exports the public surface
│   ├── anchor.py            # add_bookmark, delete_bookmark, BookmarkRef, BookmarkTarget
│   ├── crossref.py          # add_cross_reference, CrossReferenceKind
│   ├── read.py              # read_bookmarks, BookmarkInfo
│   └── registry.py          # BookmarkIdRegistry
├── notes/                   # footnotes + endnotes — v0.2
│   ├── __init__.py          # re-exports the public surface
│   ├── write.py             # add_footnote, add_endnote, edit_footnote, edit_endnote,
│   │                        # FootnoteRef, EndnoteRef, NoteNotFoundError
│   ├── read.py              # read_footnotes, read_endnotes, NoteContent
│   └── registry.py          # FootnoteIdRegistry, EndnoteIdRegistry
├── publishing/              # long-document publishing — v0.2
│   ├── __init__.py          # re-exports the public surface
│   ├── toc.py               # add_toc
│   ├── captions.py          # add_caption
│   └── figures.py           # add_table_of_figures
├── examples/                # runnable demo scripts
│   ├── inspect_document.py, restyle_existing.py, build_form.py, populate_form.py
│   ├── add_comments.py, multi_column_layout.py, bookmarks_and_xrefs.py,
│   │   footnotes_and_endnotes.py     # v0.2 demos
│   └── publishing_layout.py            # v0.2 expansion demo
└── _testing/                # internal test helpers (not public API)
    └── ooxml_asserts.py     # assert_ids_unique, assert_style_defined,
                             # count_controls, assert_protected, assert_field_dirty
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

`_BUILTIN_STYLES` in `styles/modify.py:1154` enumerates **107 of Word's
built-in styles** — well past SPEC §5's "at minimum" set, covering
essentially every style a real Word user reaches for. The entries are
grouped into seven tiers:

| Tier | Count | Coverage |
|---|---:|---|
| Core | 19 | `Normal`, `Heading1`–`Heading9`, `Title`, `Subtitle`, `Quote`, `IntenseQuote`, `ListParagraph`, `Caption`, `DefaultParagraphFont`, `Hyperlink`, `PlaceholderText`, `TableNormal`, `NoList` |
| A — structural essentials | 6 | `NoSpacing`, `Header`/`HeaderChar`, `Footer`/`FooterChar`, `TableGrid` |
| B — character emphasis | 7 | `Strong`, `Emphasis`, `IntenseEmphasis`, `SubtleEmphasis`, `IntenseReference`, `SubtleReference`, `BookTitle` |
| C — heading linked-Char | 13 | `Heading1Char`–`Heading9Char`, `TitleChar`, `SubtitleChar`, `QuoteChar`, `IntenseQuoteChar` |
| D — lists | 18 | `List`/`List2`/`List3`, `ListBullet`/`2`–`5`, `ListNumber`/`2`–`5`, `ListContinue`/`2`–`5` |
| E — TOC / index / table-of-* | 16 | `TOCHeading`, `TOC1`–`TOC9`, `IndexHeading`, `Index1`, `TableofFigures`, `TableofAuthorities`, `TOAHeading` |
| F — footnotes / endnotes / comments | 12 | `FootnoteText`/`Char`, `FootnoteReference`, `EndnoteText`/`Char`, `EndnoteReference`, `CommentText`/`Char`, `CommentReference`, `CommentSubject`/`Char`, `BalloonText`/`Char` |
| G — body / macro / preformatted | 16 | `BodyText`/`2`/`3` + Char companions, `MacroText`/`Char`, `HTMLPreformatted`/`Char`, `PlainText`/`Char`, `NormalIndent`, `BlockText` |

Defaults come from extracting `styles.xml` from real Word-saved
documents (Word 365, 2026-05-19) — *not* from guessing or copying
Word-2007 numbers. About 65 entries (Core, A, B, the most-common subset
of C–G) are sourced from python-docx's bundled `default.docx`; the
latent remainder (TOC*, footnote/endnote/comment family, Index*, table-
of-*, HTMLPreformatted, PlainText, BodyText, MacroText, BalloonText,
BlockText) were extracted from Word-saved sample docs that materialise
each style after it's applied to a paragraph.

Built-ins materialise *without* `w:customStyle="1"` (they are not
user-defined) and the four `default` entries carry `w:default="1"`.

**Known property-writer limitations.** A handful of Word's defaults
can't currently be emitted because the property writer doesn't model
them — these are intentionally omitted from `_BUILTIN_STYLES`:

- **Theme attributes** (`themeColor`, `themeShade`, `asciiTheme`, etc.)
  on `Heading*Char`, `Caption`, `IntenseQuote`, `IndexHeading`,
  `TOAHeading`. Literal RGB/font values are emitted instead — visually
  equivalent for users on Word's default Office theme.
- **`semiHidden` / `unhideWhenUsed`** presence-only flags on latent
  styles. Not a property kind we expose; styles still work, they just
  always show in Word's style gallery.
- **Tab stops** on `Header`, `Footer`, `MacroText`.
- **Paragraph borders** (`pBdr`) on `IntenseQuote`, `BlockText`.
- **`numPr` placeholder** on `ListBullet`/`ListNumber` — these styles
  in Word's default ship with an empty `numPr` child (a hint, no real
  numbering link). Skipped; callers attach numbering separately.

`ensure_style` is idempotent and aware that **python-docx already ships
a `styles.xml` with many of these latent built-ins materialised** at
Word-2007 defaults (e.g. Heading1 = 14pt #365F91), not Word-2013/365.
This is deliberate: `ensure_style` consults the built-ins table **only**
when the ID is genuinely missing from `styles.xml`. If python-docx
already shipped it, the existing definition is returned unchanged. The
table is a "the style is absent, here is what Word would have written"
fallback, not a "force my preferred defaults" mechanism — for that,
use `modify_style` or `remap_styles`.

---

## §6 Content controls

`controls/builder.py:FormBuilder` is the build-side surface and
`controls/read.py` is the read/modify side. Both target the five SDT
control types Word's UI ribbon offers: text (single- and multi-line),
dropdown / combobox, date picker, and checkbox. Rich-text SDTs (no
marker child) are recognised but skipped — they're a v0.2 deferred case.

### `FormBuilder`

The wrapper accepts an existing `Document`, a path, or `None` (start
fresh). On construction it does three things:

1. **Materialises the `PlaceholderText` character style** in
   `styles.xml` if it's absent — without it Word's grey placeholder
   text fails to render. This duplicates the style definition rather
   than importing it from `styles/modify.py` (SPEC §9.1 forbids
   capability-to-capability imports).
2. **Verifies the `w14` namespace is declared on the document root.**
   Required by `w14:checkbox`. python-docx 1.2.0 declares it by default;
   if a future version drops it, construction raises `MissingNamespaceError`.
3. **Seeds an `IdRegistry`** from existing SDT IDs in the body, or
   accepts one passed in via the `id_registry=` kwarg for callers that
   need to share allocation across multiple builders.

Each `add_*` method appends its SDT inline at the end of the paragraph
you pass — so put the field's label text in the paragraph first. The
SDT's `w:sdtPr` children are emitted in CT_SdtPr schema order
(`alias? → tag → id → showingPlcHdr? → <type-marker>`). The `<type-marker>`
distinguishes the controls: `w:text` for text/multiline, `w:dropDownList`
or `w:comboBox` for selectors, `w:date` for date pickers, `w14:checkbox`
for checkboxes.

### `read_controls` and `set_control_value`

`read_controls(doc, *, by="tag")` returns a `dict[str, ControlValue]`
keyed by tag (default) or alias. Control-type dispatch lives in
`_classify_sdt` and is shared with `_testing.ooxml_asserts.count_controls`
so there is one source of truth. Repeating tags raise `DuplicateTagError`
— a precondition v0.1 enforces because Custom-XML-Part data binding
(the v0.2 feature that supports repeating sections) isn't shipped yet.

`set_control_value(doc, tag, value)` accepts `str | bool | datetime`
matched against the control type. Type mismatches raise
`ControlTypeError`. Dropdowns try `w:value` first then `w:displayText`,
raising `ValueNotInListError` if neither matches — unless the control
is a combobox, in which case any string is accepted (matching Word's
freeform-input behaviour). Date values round-trip through
`w:date/@w:fullDate` (ISO 8601); the human-readable rendered text in
`sdtContent` is best-effort because full Word date-format-token
translation is a v0.2 concern.

`clear_control(doc, tag)` resets to the placeholder state.

---

## §7 Fields and protection

`fields/` covers complex-field insertion and the "Word recalculates on
open" flag; `protection/` covers document-level enforcement. Both are
small modules (≤100 lines each) and mostly schema-strict insertion into
`settings.xml`.

### Complex fields

A Word field is **not** a single element. It's a sequence of five runs
that bracket an instruction (`w:instrText`) and a cached result (`w:t`):

```
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t xml:space="preserve">1</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

`core/oxml.py:build_complex_field` (hoisted from
`fields/simple.py` in v0.2 so cross-references can reuse it without a
cross-capability import) is the single helper that emits this sequence.
`fields/simple.py`'s three public functions (`add_page_number_field`,
`add_date_field`, `add_field`) all route through it, as does
`bookmarks/crossref.py:add_cross_reference`. Both the instruction and
the cached result carry `xml:space="preserve"` so Word's XML reader does
not collapse the spaces that the field-instruction grammar requires.

Each public helper returns the begin `<w:r>` element so callers can
navigate or relocate the field. The `xml` namespace was added to
`core/ns.py:NSMAP` in Phase 5 specifically to make `qn("xml:space")`
work; before that the prefix was unknown to the library.

### `mark_fields_dirty`

`fields/update.py:mark_fields_dirty(doc)` writes
`<w:updateFields w:val="true"/>` into `settings.xml`. Word reads this
flag on open, recalculates every field in the document, and resets the
flag to `false` — it's a one-shot mechanism, not persistent state. The
function is idempotent: a second call updates the existing element
rather than duplicating it.

### `protect_document`

`protection/document.py:protect_document(doc, *, mode=...)` emits
`<w:documentProtection w:edit="MODE" w:enforcement="1"/>` into
`settings.xml`. `mode` accepts the four `ProtectionMode` literals:

- `"forms"` (default) — only content controls are editable. Pair with
  `FormBuilder` to produce a fillable form readers can't drift outside.
- `"readOnly"` — entire document is read-only.
- `"comments"` — readers may only add comments.
- `"trackedChanges"` — readers may edit with revisions on.

Idempotent: a second call replaces the mode rather than stacking.
`unprotect_document(doc)` removes the element, no-op when absent.
`is_protected(doc)` is the presence predicate (does not introspect the
mode).

Protection is **unpassworded** in v0.1 (SPEC §1 non-goal). The
`w:enforcement="1"` flag stops accidental editing in Word's UI but does
not stop a determined user from rewriting `settings.xml`.
Password-protected forms (legacy hash algorithm) are deferred to v0.2.

### Schema-strict insertion in `settings.xml`

`w:documentProtection`, `w:updateFields`, and (v0.2) `w:evenAndOddHeaders`
all live deep in `CT_Settings`'s child sequence (ECMA-376 17.15.1.78).
Inserting them at the wrong position produces a file Word will silently
"repair" on open — sometimes correctly, sometimes not. Every callsite
applies the same `core/oxml.py:insert_before_first_anchor(parent,
new_element, anchor_tags)` pattern, walking a tuple of later-siblings
(`w:defaultTabStop`, `w:compat`, `w:rsids`, etc.) and inserting before
the first match. If no anchor is present, the helper falls back to
appending — the no-anchor case is exercised by
`test_mark_fields_dirty_appends_when_no_anchor`. The helper lives in
`core/oxml.py` (hoisted in v0.2 when `layout/settings.py` became the
third caller); the per-module anchor tuples stay co-located with their
callsites so the schema position is reviewed alongside the new child.

---

## §7.5 Separate OOXML parts

v0.1 capabilities (styles, fields, controls, protection) only mutated
the main document part and `settings.xml`. v0.2 introduces three
capabilities backed by **separate** parts that may not exist in a
fresh document:

- `/word/comments.xml` (relationship `RT.COMMENTS`)
- `/word/footnotes.xml` (relationship `RT.FOOTNOTES`)
- `/word/endnotes.xml` (relationship `RT.ENDNOTES`)

`core/parts.py:get_or_create_part(doc, spec)` is the single entry
point. Given a `PartSpec` describing the target, it tries
`doc.part.part_related_by(spec.relationship_type)`; on `KeyError` it
parses `spec.root_xml` for the empty default root element, looks up
the correct part class from `PartFactory.part_type_for`, constructs the
part, and wires the relationship. Returns `(part, root_element)`.

python-docx already registers `CommentsPart` for `WML_COMMENTS` at
package-import time. It does **not** register footnote or endnote
content types, so `core/parts.py` does — installing internal
`_FootnotesPart` / `_EndnotesPart` subclasses of `XmlPart` with
`PartFactory.part_type_for.setdefault(...)`. Without that registration,
an existing document with footnotes would deserialize the part as the
default `Part` (blob-only), and `part.element` would not exist.

Three pre-baked `PartSpec` constants cover every v0.2 need:
`COMMENTS_SPEC`, `FOOTNOTES_SPEC`, `ENDNOTES_SPEC`. Custom callers can
build their own.

---

## §7.6 Anchored comments

`comments/anchor.py:add_comment(target, text, ...)` is the v0.2
headline. Closes the largest python-docx gap: python-docx 1.x writes
`<w:comment>` into `comments.xml` but skips the three body-side
elements that anchor the comment to a text range, so its comments show
in the review pane but have nothing to point at when the reader clicks
"show in document".

Each `add_comment` writes four elements:

1. `<w:commentRangeStart w:id=N/>` — placed before `start_anchor` via
   `addprevious`
2. `<w:commentRangeEnd w:id=N/>` — placed after `end_anchor` via
   `addnext`
3. The reference run — `<w:r><w:rPr><w:rStyle val="CommentReference"/></w:rPr><w:commentReference w:id=N/></w:r>`
   — placed after the range end
4. The comment body — `<w:comment w:id=N w:author=... w:date=... [w:initials=...]>`
   appended to the root of `comments.xml` (via `get_or_create_part`)

Target shapes: a python-docx `Run` (brackets just that run), a
`Paragraph` (brackets every run, must have ≥1 run), or a
`(start_run, end_run)` tuple for a range. Range tuples may span
paragraphs; OOXML permits this. Comment body uses
`xml:space="preserve"` so leading/trailing whitespace survives Word's
XML reader.

`delete_comment(doc, comment_id)` is the inverse — removes all four
elements and is idempotent (missing id is a no-op).

`read_comments(doc)` walks `comments.xml` and pairs each `<w:comment>`
with its body range, extracting `author`, `initials`, `timestamp`
(parsed `xsd:dateTime`), the comment `text`, the `anchored_text`
between the body markers, and the `paragraph_index` where the
`commentRangeStart` sits. Orphaned comments (no matching body range)
appear with `anchored_text=""` and `paragraph_index=-1`.

`CommentIdRegistry` lives in its own namespace (separate from SDT,
bookmark, note ids). It seeds from both the comments part AND any
orphaned body-side anchors so a partially-deleted comment cannot
trigger id reuse.

Threaded comments (w15 `parentCommentEx` for replies) are deferred to
v0.3 — basic anchored comments first.

---

## §7.7 Layout

`layout/` ships five documented python-docx gaps. None of them
duplicate functionality python-docx already exposes (orientation,
margins, page size, per-section header / footer, `add_section`).

**`set_columns(section, num, *, space, separator, widths)`** in
`layout/columns.py` emits `<w:cols w:num=... w:space=... w:sep=...>`
into the section's `sectPr`. Idempotent — replaces any existing
`<w:cols>`. With `widths` supplied, it emits per-column `<w:col>`
children with `w:equalWidth="0"` so Word reads widths from the children
rather than the parent `w:space`.

**`insert_section_break(paragraph, *, start_type)`** in
`layout/breaks.py` handles the case `Document.add_section` does not —
inserting a break mid-document. The algorithm clones the trailing
body-level `<w:sectPr>` (the document's "sentinel"), sets `<w:type>`
on the clone to the requested start kind, and calls python-docx's
`CT_P.set_sectPr(clone)` to embed it in the chosen paragraph's `pPr`.
The new section inherits all properties (page size, margins, header /
footer references) from the sentinel; both sections render with the
same headers and footers unless the caller mutates the returned
`Section` proxy.

**`enable_distinct_even_odd_headers(doc)`** in `layout/settings.py`
writes `<w:evenAndOddHeaders/>` into `settings.xml` via the
schema-strict insertion pattern (§3). This flag is constantly confused
with two other things: the per-section `<w:titlePg>` (controls whether
*first* page has a distinct header/footer, exposed by python-docx as
`Section.different_first_page_header_footer`), and the per-section
header/footer reference types (`w:headerReference w:type="even"`,
which Word reads *because* the doc-level flag is set). All three are
required for a real even-page-distinct workflow. `disable_…` removes
the doc-level element; both functions are idempotent.

**`set_line_numbering(section, *, count_by, restart, start, distance)`**
in `layout/line_numbering.py` emits `<w:lnNumType>` into the section's
`sectPr` — Word's mechanism for the marginal line numbers that legal
and contract documents require. Schema-strict via
`core.insert_before_first_anchor`; the element lands in its
ECMA-376 17.6.17 slot regardless of which other `sectPr` children
exist. `restart` is the only argument that validates eagerly (one of
`"newPage"` / `"newSection"` / `"continuous"`); `count_by` and `start`
must be ≥ 1. Idempotent.

**`set_page_borders(section, *, top, bottom, left, right)`** in
`layout/borders.py` emits `<w:pgBorders>` from a `Border` dataclass
per side (`style`, `size` in eighths of a point, `color`, `space` in
twips). Sides set to `None` are omitted from the emitted XML; passing
all four as `None` removes the element rather than emitting an empty
container. Schema-strict, idempotent.

---

## §7.8 Bookmarks and cross-references

`bookmarks/anchor.py:add_bookmark(target, name, ...)` writes a paired
`<w:bookmarkStart w:id=N w:name=...>` / `<w:bookmarkEnd w:id=N/>`
around the target. Target shapes mirror `add_comment`: `Run`,
`Paragraph` (≥1 run), or `(Run, Run)` tuple. The name is validated
against Word's bookmark rules: `[A-Za-z_][A-Za-z0-9_]{0,39}`. Words
with spaces or punctuation are silently rejected by Word's UI but
accepted in raw OOXML, which leads to confusing failures —
`add_bookmark` raises eagerly instead.

`delete_bookmark(doc, name)` removes every bookmark with the given
name (by name, not id, because that's what cross-references key off).
`read_bookmarks(doc)` returns a `BookmarkInfo` per bookmark with id,
name, anchored text, and paragraph index. `BookmarkIdRegistry` is the
fourth namespace (after SDT, comment, footnote / endnote each get
their own).

`bookmarks/crossref.py:add_cross_reference(paragraph, *, bookmark,
kind, hyperlink)` builds a `REF` (`kind="text"`) or `PAGEREF`
(`kind="page"`) complex field via `core.build_complex_field`. The
`\h` flag is appended by default so Word renders the cross-reference
as a clickable link. Pair calls with `mark_fields_dirty` so Word
recalculates the cached results on first open.

Cross-references to headings, numbered items, or captions
(`STYLEREF`, sequence fields) are deferred to v0.3 — they require
different field instructions but the same field-building plumbing.

---

## §7.9 Footnotes and endnotes

`notes/write.py` exposes `add_footnote` and `add_endnote`, both with
identical shape: append a reference marker run to the paragraph, then
append a content entry in the corresponding separate part. The content
entry uses Word's `FootnoteText` / `EndnoteText` paragraph style and
`FootnoteReference` / `EndnoteReference` run style for the leading
reference glyph. The body text run carries `xml:space="preserve"`.

`edit_footnote(doc, id, text)` and `edit_endnote(doc, id, text)` mutate
the body of an existing note in place. They strip every `<w:p>` child
of the matching `<w:footnote>` / `<w:endnote>` element and append a
fresh paragraph built by the shared `_build_note_paragraph` helper
(used by both add and edit paths). The body-side reference marker in
the main document body is untouched, so the in-text superscript stays
put. Reserved separator ids (`-1`, `0`) raise `ValueError`; missing
ids raise `NoteNotFoundError`.

`read_footnotes(doc)` and `read_endnotes(doc)` walk the corresponding
part and pair each note with the paragraph index of its body-side
reference marker. Reserved entries (ids `-1` for separator, `0` for
continuation separator, or any entry with `w:type` of `"separator"` /
`"continuationSeparator"`) are filtered out before results are
returned, so callers only ever see user-authored notes.

`FootnoteIdRegistry` and `EndnoteIdRegistry` are two more disjoint
namespaces. The shared `_NoteIdRegistryBase` (`notes/registry.py`)
parameterises the relationship type and the note tag; the underlying
`_IdRegistryBase.reserve(value)` rejects values outside `[1, 2**31 - 1]`
on a range check, so ids `0` and `-1` are unissuable — the range check
fires before any duplicate check, so no special pre-seeding is
needed.

---

## §7.10 Publishing

`publishing/` composes the existing fields plumbing into the
long-document primitives that make Word a viable publishing target.
Three helpers, each emitting a single complex field on top of
`core.build_complex_field`:

- `add_toc(paragraph, *, levels=(1, 3), hyperlink=True, page_numbers=True)`
  emits a `TOC` field. The instruction string is assembled from
  kwargs: `\o "lo-hi"` for outline-level range, `\h` for hyperlinked
  entries, the always-present `\z` and `\u` (Word emits both by
  default), and the optional `\n` to suppress page numbers.
- `add_caption(paragraph, label, *, caption_type="Figure", numbering="ARABIC")`
  emits a label text run (`"Figure "`) followed by a `SEQ` complex
  field. Items sharing the same `caption_type` auto-number together;
  the name is the same vocabulary a Table of Figures uses via its
  `\c` switch.
- `add_table_of_figures(paragraph, *, caption_type="Figure", hyperlink=True)`
  emits `TOC \c "<caption_type>"`, structurally a TOC keyed off the
  matching SEQ captions instead of paragraph outline levels.

None of the three auto-calls `mark_fields_dirty`. The publishing
module respects the §8 invariant of importing only from `core/`, and
forwarding to `fields/` would violate it. Users pair their
publishing inserts with one explicit `mark_fields_dirty(doc)` call
before save — the docstrings document the contract.

Bibliography (sources stored in a Custom XML Part, `<w:sdt>`
citations referencing them, a `BIBLIOGRAPHY` field rendering the
list) is deferred to v0.3 because it depends on the CXML
data-binding subsystem (also v0.3).

---

## §8 Invariants

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
   the stdlib bases. See §9.

8. **No unrequested side effects on the input document.** Functions
   that mutate document state document the mutation in the docstring.
   `resolve_*` and `read_*` functions are pure reads.

---

## §9 Error hierarchy

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
| `MissingPartError` | `DocxPlusError` | `styles/inspect.py` | A referenced part is required but absent (currently unused — see §2 layer 4) |
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

## §10 Testing strategy

SPEC §10 specifies three layers:

- **Layer 1 — structural unit tests.** One file per module, fast, no
  I/O beyond reading fixtures. **631 tests** at the v0.2.0 release:
  v0.1's surface (319 tests) plus the v0.2 cycle — `core/parts` (13),
  `comments/` (35), `layout/` (47), `bookmarks/` + cross-refs (26),
  `notes/` (34), `styles/` table conditional (13), `publishing/` (23)
  — plus example smoke tests for the new demos, plus the regression
  coverage added by the pre-publication code/docs review (cascade
  correctness, schema/part wiring, error taxonomy, publishing
  validation, and the six newly-writable run toggles).
- **Layer 2 — round-trip tests.** Build → save → reopen with
  `python-docx` → assert. The high-value class for OOXML
  correctness (`IMPLEMENTATION.md §8`). Phase 5 added round-trips for
  every field type plus the protect/unprotect cycle;
  `TEST_GAPS.md` I1 lists the remaining gaps on the modify side.
- **Layer 3 — headless render smoke.** Run each example, convert to
  PDF with LibreOffice headless, assert exit-0 and page count. Gated
  on the `requires_libreoffice` pytest marker; deferred to Phase 6.

Test fixtures live in `tests/fixtures/build_fixtures.py` (the build
script is the source of truth, not the `.docx` files it produces —
`.gitignore` excludes the generated docx files). `empty.docx`,
`multistyle.docx`, `themed.docx`, and `existing_form.docx` are built
on demand.

Shared assertions live in `docx_plus/_testing/ooxml_asserts.py`:
`assert_ids_unique`, `assert_style_defined`, `count_controls`,
`assert_protected`, `assert_field_dirty`. The module is internal —
not re-exported from the top-level package — and is built out lazily
as later tests demand more helpers. Of the SPEC §10 helper list, only
`assert_style_not_defined` and `assert_no_orphan_relationships`
remain unwritten.

For a frozen snapshot of where the suite has real holes, see
[`TEST_GAPS.md`](TEST_GAPS.md).

---

## §11 What's next

v0.1 (Phases 1–6), the v0.2 cycle, and the v0.2 in-place expansion
(scoped in SPEC §15) are complete. The pieces deferred to v0.3+ are:

- **Threaded comments** (w15 `parentCommentEx` for parent/child
  replies) and the **respond / resolve / reopen** ops that depend on
  them. Adds a `w15` namespace dependency and a separate
  `commentsExtended.xml` part.
- **Cross-references to non-bookmark targets** — `STYLEREF` for
  heading-text references, sequence fields for caption / figure
  numbering. Reuses the same complex-field plumbing; the work is the
  instruction grammar.
- **CLI** — `docx-plus restyle` for style remapping plus `inspect`
  (dump effective formatting) and `controls` (list / set values)
  subcommands.
- **Custom XML Parts data binding** — wires repeating-section content
  controls to a custom XML data source. The plumbing in
  `core/parts.py` already supports separate parts; the binding adds
  relationship types and `<w:dataBinding>` children on SDTs.
- **Bibliography** — sources stored in a Custom XML Part, `<w:sdt>`
  citations referencing them, a `BIBLIOGRAPHY` field rendering the
  list. Rides on the CXML data-binding subsystem above.
- **Tracked changes** — read/write API for the OOXML revision marks
  (`w:ins`, `w:del`, `w:moveFromRangeStart`, etc.). Significant scope.
- **Glossary placeholder text** — the "formal" placeholder mechanism
  for SDTs (vs. the inline `w:placeholder` text the controls module
  already supports).
- **Password-protected forms** — legacy hash algorithm, paired with
  `protect_document`.
- **Theme writing** — `styles/theme.py` reads themes today; writing
  rounds out the surface.
- **A high-level "restyle" planner** — inverse of the inspector, takes
  a target `ResolvedFormatting` and computes the minimal cascade
  modification to reach it.
- **Sections / headers / footers first-class API** — wraps the
  python-docx primitives behind a docx_plus-native surface.

Everything in this list is enumerated in `SPEC.md §15`. The roadmap
order is driven by `notes-*.md` discussion artifacts at the repo
root, which capture user-flagged wants and inform priority.
