# Fields and protection

`fields/` covers complex-field insertion and the "Word recalculates on
open" flag; `protection/` covers document-level enforcement. Both are
small modules (≤100 lines each) and mostly [schema-strict
insertion](schema-order.md) into `settings.xml`.

For the calls, see the [fields guide](../guides/fields.md) and the
[forms guide](../guides/forms.md).

## Complex fields

A Word field is **not** a single element. It's a sequence of five runs
that bracket an instruction (`w:instrText`) and a cached result (`w:t`):

```xml
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

## `mark_fields_dirty`

`fields/update.py:mark_fields_dirty(doc)` writes
`<w:updateFields w:val="true"/>` into `settings.xml`. Word reads this
flag on open, recalculates every field in the document, and resets the
flag to `false` — it's a one-shot mechanism, not persistent state. The
function is idempotent: a second call updates the existing element
rather than duplicating it.

This is the single most commonly forgotten call in the library. Every
surface that emits a field — [publishing](publishing.md),
[cross-references](bookmarks.md), page numbers, dates — produces a blank
result on disk until Word recalculates it.

## `protect_document`

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
Password-protected forms (legacy hash algorithm) are on the backlog.
