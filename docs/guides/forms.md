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
from docx_plus.controls import clear_control, read_controls, set_control_value

doc = Document("form.docx")

for tag, ctrl in read_controls(doc).items():
    print(tag, ctrl.control_type, repr(ctrl.value), ctrl.is_placeholder)

set_control_value(doc, "full_name", "Ada Lovelace")           # text     -> str
set_control_value(doc, "dept", "ENG")                         # dropdown -> str
set_control_value(doc, "remote", True)                        # checkbox -> bool
set_control_value(doc, "start_date", datetime(2026, 6, 1))    # date     -> datetime

clear_control(doc, "notes")      # back to the placeholder state

doc.save("form_filled.docx")
```

`read_controls(doc, *, by="tag")` returns a `dict[str, ControlValue]` — key
by `"alias"` instead if you prefer. A `ControlValue` is a frozen dataclass
with `tag`, `alias`, `control_type`, `value`, and `is_placeholder`.
`control_type` is one of `"text"`, `"dropdown"`, `"combobox"`, `"date"`,
`"checkbox"`.

For a closed dropdown, `set_control_value` accepts either the stored value
or the visible display text; anything else raises `ValueNotInListError`.
Comboboxes accept free-form input and never raise it.

!!! note "Tags must be unique"
    `read_controls` raises `DuplicateTagError` if two controls share a tag,
    because the returned dict would be ambiguous. Repeating sections need
    Custom XML Part data binding, which is on the backlog.

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
| `ControlNotFoundError` (`KeyError`) | `set_control_value` / `clear_control` on an unknown tag |
| `DuplicateTagError` (`ValueError`) | Two controls share a tag |
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
