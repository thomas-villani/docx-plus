# Forms and protection

Build fillable Word forms with **content controls** (Structured Document
Tags, or SDTs), read and write their values, and lock the surrounding text
so only the controls accept input.

Two modules: `docx_plus.controls` for build and read/write,
`docx_plus.protection` for the document-level lock.

## Building a form

`FormBuilder` is the one place `docx_plus` wraps a document. It seeds an id
registry, materialises the `PlaceholderText` style (without which Word's
grey placeholder text doesn't render), and verifies the `w14` namespace
checkboxes need.

```python
from docx_plus.controls import FormBuilder

fb = FormBuilder()                    # blank document
# fb = FormBuilder("template.docx")   # or open a path
# fb = FormBuilder(existing_doc)      # or wrap an open Document

fb.doc.add_heading("New employee form", level=1)   # fb.doc is real python-docx
```

Add ordinary content through `fb.doc`; add controls through the builder
methods. Every control attaches **inline to a paragraph you pass in**, so
put the field's label text in the paragraph first.

Each control takes a `tag=` (the machine key you use later to read and set
the value) and an optional `alias=` (the human label Word shows in the
control's title bar).

### Text

```python
p = fb.doc.add_paragraph("Full name: ")
fb.add_text_control(p, tag="full_name", alias="Full name",
                    placeholder="Type your name")

p = fb.doc.add_paragraph("Notes: ")
fb.add_text_control(p, tag="notes", placeholder="(optional)", multiline=True)
```

`add_text_control(paragraph, *, tag, alias=None,
placeholder="Click to enter text", multiline=False)`

### Dropdown and combobox

`items` is either a list of display strings, or a list of
`(display, value)` tuples when the stored value differs from what's shown.
A plain dropdown is closed; `editable=True` makes it a **combobox** that
also accepts free-form typing.

```python
p = fb.doc.add_paragraph("Department: ")
fb.add_dropdown(p, tag="dept",
                items=[("Engineering", "ENG"), ("Design", "DES"), ("Ops", "OPS")],
                placeholder="Choose a department")

p = fb.doc.add_paragraph("Office: ")
fb.add_dropdown(p, tag="office", items=["New York", "London", "Remote"],
                editable=True)      # combobox — free-form allowed
```

`add_dropdown(paragraph, *, tag, items, alias=None,
placeholder="Choose an item", editable=False)`

### Date picker

`date_format` is a Word date pattern (`"M/d/yyyy"`, `"MMMM d, yyyy"`, …).

```python
p = fb.doc.add_paragraph("Start date: ")
fb.add_date_picker(p, tag="start_date", date_format="M/d/yyyy")
```

`add_date_picker(paragraph, *, tag, alias=None,
placeholder="Click to select a date", date_format="M/d/yyyy",
lcid="en-US")`

### Checkbox

```python
p = fb.doc.add_paragraph("Remote-first? ")
fb.add_checkbox(p, tag="remote", checked=False)
```

`add_checkbox(paragraph, *, tag, alias=None, checked=False)`

Checkboxes need the `w14` namespace on the document root. `FormBuilder`
guarantees it at construction, raising `MissingNamespaceError` if a future
python-docx stops declaring it.

### Save

```python
fb.save("form.docx")     # returns the path as str
```

## Reading and writing values

These operate on a plain python-docx `Document`, not on `FormBuilder` — so
you can fill a form you never built.

```python
from datetime import datetime

from docx import Document
from docx_plus.controls import clear_control, list_controls, set_control_value

doc = Document("form.docx")

for ctrl in list_controls(doc):
    print(ctrl.tag, ctrl.control_type, repr(ctrl.value), ctrl.is_placeholder)

set_control_value(doc, "full_name", "Ada Lovelace")           # text     -> str
set_control_value(doc, "dept", "ENG")                         # dropdown -> str
set_control_value(doc, "remote", True)                        # checkbox -> bool
set_control_value(doc, "start_date", datetime(2026, 6, 1))    # date     -> datetime

clear_control(doc, "notes")      # back to the placeholder state

doc.save("form_filled.docx")
```

`list_controls(doc)` returns a `list[ControlValue]` in document order — every
control, nothing keyed, nothing dropped. A `ControlValue` is a frozen dataclass
with `tag`, `alias`, `control_type`, `value`, `is_placeholder`, `control_id`,
`index`, and `location`. `control_type` is `"text"`, `"dropdown"`,
`"combobox"`, `"date"`, or `"checkbox"` for the controls you can set a value
on, and one of `"richtext"`, `"picture"`, `"group"`, `"repeating"`,
`"repeatingitem"`, `"docpart"`, `"citation"`, `"bibliography"`, `"equation"`
for the ones that hold block content instead.

`read_controls(doc, *, by="tag")` is the keyed convenience on top of it,
returning a `dict[str, ControlValue]` — key by `"alias"` instead if you prefer.

For a closed dropdown, `set_control_value` accepts either the stored value
or the visible display text; anything else raises `ValueNotInListError`.
Comboboxes accept free-form input and never raise it.

!!! warning "Tags are neither required nor unique"
    This is the one thing to know before pointing this API at a document Word
    produced rather than one `FormBuilder` built.

    A control inserted from Word's Developer ribbon is written with
    `<w:tag w:val=""/>` unless the author opens the properties dialog and types
    a tag. Most never do, so a real form usually has *one empty tag shared by
    every control* — and sometimes no `w:tag` element at all.

    `read_controls` can only report controls that have a usable key, so on
    such a document it returns almost nothing. **Use `list_controls`**, and
    address individual controls by `control_id` (the `w:id`, which is what
    OOXML actually uses for identity):

    ```python
    for ctrl in list_controls(doc):
        if ctrl.alias == "Client name":
            set_control_value(doc, None, "Acme Corp", control_id=ctrl.control_id)
    ```

    `read_controls` still raises `DuplicateTagError` when two controls share a
    *non-empty* key, and so do `set_control_value` / `clear_control` when a tag
    matches more than one control — writing to an arbitrary match would leave
    the others untouched while reporting success. `control_id` is the way past
    it.

!!! note "Controls outside the body"
    Both read functions also walk headers, footers, footnotes, and endnotes,
    reporting the story in `ControlValue.location` (`"body"`,
    `"header:1:primary"`, `"footnotes"`, …).

## Locking the document

```python
from docx_plus.protection import is_protected, protect_document, unprotect_document

protect_document(doc, mode="forms")   # only content controls are editable

is_protected(doc)         # -> bool (presence check; doesn't report the mode)
unprotect_document(doc)   # idempotent
```

`mode` is one of:

| Mode | Effect |
|---|---|
| `"forms"` (default) | Only content controls are editable — the form case |
| `"readOnly"` | The whole document is read-only |
| `"comments"` | Readers may only add comments |
| `"trackedChanges"` | Readers may edit, with revisions forced on |

!!! warning "This is a UI guard, not encryption"
    Protection is **unpassworded**. The `w:enforcement="1"` flag stops
    accidental editing in Word's UI but does not stop anyone who is willing
    to rewrite `settings.xml`.

## End to end

```python
from datetime import datetime

from docx import Document
from docx_plus.controls import FormBuilder, read_controls, set_control_value
from docx_plus.protection import protect_document

# --- build ---
fb = FormBuilder()
fb.doc.add_heading("Onboarding", level=1)
p = fb.doc.add_paragraph("Name: ");    fb.add_text_control(p, tag="name")
p = fb.doc.add_paragraph("Team: ");    fb.add_dropdown(p, tag="team",
        items=[("Engineering", "ENG"), ("Design", "DES")])
p = fb.doc.add_paragraph("Start: ");   fb.add_date_picker(p, tag="start")
p = fb.doc.add_paragraph("Remote? ");  fb.add_checkbox(p, tag="remote")
protect_document(fb.doc, mode="forms")
fb.save("form.docx")

# --- fill ---
doc = Document("form.docx")
set_control_value(doc, "name", "Ada Lovelace")
set_control_value(doc, "team", "ENG")
set_control_value(doc, "start", datetime(2026, 6, 1))
set_control_value(doc, "remote", True)
doc.save("form_filled.docx")

# --- read back ---
for tag, c in read_controls(Document("form_filled.docx")).items():
    print(f"{tag}: {c.value!r}")
```

The same fill-and-read job from a shell is [`docx-plus
controls`](../cli.md).

## Errors

All subclass `DocxPlusError`, and most the noted builtin.

| Error | Raised when |
|---|---|
| `ControlNotFoundError` (`KeyError`) | `set_control_value` / `clear_control` on an unknown tag or `control_id`, or with neither given |
| `DuplicateTagError` (`ValueError`) | A tag doesn't identify exactly one control — two share a non-empty key on read, or a writer's tag matched several. An absent or empty tag is unkeyable rather than duplicate |
| `ValueNotInListError` (`ValueError`) | A closed-dropdown value matches no item |
| `ControlTypeError` (`TypeError`) | The value's type doesn't match the control type |
| `MissingNamespaceError` | `add_checkbox` on a doc without `w14` declared |
| `InvalidDropdownItemError` (`TypeError`) | An `items` entry is neither `str` nor a `(display, value)` tuple |

## See also

- [How content controls work](../concepts/controls.md) and [fields and
  protection](../concepts/fields.md)
- Reference: [`controls.builder`](../reference/controls-builder.md),
  [`controls.read`](../reference/controls-read.md),
  [`protection.document`](../reference/protection-document.md)
- Examples: `build_form.py`, `populate_form.py`
