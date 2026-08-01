# Content controls

`controls/builder.py:FormBuilder` is the build-side surface and
`controls/read.py` is the read/modify side. Both target the five SDT
control types Word's UI ribbon offers: text (single- and multi-line),
dropdown / combobox, date picker, and checkbox. Rich-text SDTs (no
marker child) are recognised but skipped — they're a v0.2 deferred case.

For the calls, see the [forms guide](../guides/forms.md).

## `FormBuilder`

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

## `read_controls` and `set_control_value`

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
