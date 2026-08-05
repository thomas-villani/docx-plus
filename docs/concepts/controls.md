# Content controls

`controls/builder.py:FormBuilder` is the build-side surface and
`controls/read.py` is the read/modify side.

The **write** surface targets the five SDT control types Word's UI ribbon
offers a scalar value for: text (single- and multi-line), dropdown / combobox,
date picker, and checkbox. These are `WRITABLE_TYPES`.

The **read** surface is wider, and deliberately so: it reports every `w:sdt`,
including rich-text and the container types (`group`, `repeatingSection`,
`docPartObj`, `picture`, `citation`, `bibliography`, `equation`), which hold
block-level content rather than a value. Anything the read side cannot classify
would otherwise be invisible, and a control that silently does not exist is a
worse failure than one that reports itself as read-only.

For the calls, see the [forms guide](../guides/forms.md).

## `w:tag` is not a primary key

This is the assumption that most often breaks code written against
`FormBuilder` output and then pointed at a real Word document.

`FormBuilder` gives every control a deliberate, unique tag. OOXML does not
require either property, and Word does not supply them: a control inserted from
the Developer ribbon is written with `<w:tag w:val=""/>` unless the author
opens the properties dialog and types a tag, which most authors never do. A
real form therefore tends to have one empty tag shared across every control,
and sometimes no `w:tag` element at all.

`ControlValue` distinguishes the three states — a tag string, `""` for a
present-but-empty tag, and `None` for no tag element — and carries
`control_id` (the `w:id`, OOXML's actual identity field) for addressing.
Aliases are no better: they are UI labels and repeat freely.

The practical consequence is the split between the two read functions:

| | `list_controls(doc)` | `read_controls(doc, by=...)` |
| --- | --- | --- |
| Returns | `list[ControlValue]`, document order | `dict[str, ControlValue]` |
| Untagged / empty-tag controls | included | omitted (no usable key) |
| Repeated key | fine — no keying | `DuplicateTagError` |
| Use it for | any document, especially Word-authored | forms you built and tagged yourself |

## Which story a control lives in

Controls are not confined to the body. Both read functions walk the body, every
section's explicitly-defined headers and footers, and the footnote and endnote
parts, reporting the story in `ControlValue.location`. The traversal skips
header/footer slots whose `is_linked_to_previous` is true — partly to avoid
reporting an inherited definition once per section that inherits it, and partly
because python-docx *creates* the part on first access to an undefined one,
which would make a read mutate the document.

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

## `list_controls`, `read_controls`, and `set_control_value`

`list_controls(doc)` is the primitive: one `ControlValue` per `w:sdt`, in
document order, keying nothing. `read_controls(doc, *, by="tag")` is built on
it and returns a `dict[str, ControlValue]` keyed by tag (default) or alias,
skipping controls whose key is absent or empty and raising `DuplicateTagError`
on a genuinely repeated key.

Control-type dispatch lives in `_classify_sdt` and is shared with
`_testing.ooxml_asserts.count_controls` so there is one source of truth. A
`w:sdt` with no marker child classifies as `richtext` rather than being
skipped: ECMA-376 §17.5.2 makes rich text the default for that choice group,
which is why Word omits the marker on the control it inserts most often.

`set_control_value(doc, tag, value, *, control_id=None)` accepts
`str | bool | datetime` matched against the control type. Type mismatches
raise `ControlTypeError`, as does a control outside `WRITABLE_TYPES`.
Dropdowns try `w:value` first then `w:displayText`, raising
`ValueNotInListError` if neither matches — unless the control is a combobox,
in which case any string is accepted (matching Word's freeform-input
behaviour). Date values round-trip through `w:date/@w:fullDate` (ISO 8601);
the human-readable rendered text in `sdtContent` is best-effort because full
Word date-format-token translation is a v0.2 concern.

A `tag` that matches more than one control raises `DuplicateTagError` rather
than writing to the first match, which would leave the other matches untouched
while reporting success. Pass `control_id=` — from
`ControlValue.control_id` — to target one unambiguously; it takes precedence
over `tag`, so `tag` may be `None`.

`clear_control(doc, tag, *, control_id=None)` resets to the placeholder state
and follows the same selection rules.
