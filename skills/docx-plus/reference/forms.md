# Forms — content controls & protection

Build fillable Word forms with **content controls** (Structured Document Tags /
SDTs), read and write their values, and lock the surrounding text so only the
controls accept input.

Modules: `docx_plus.controls` (build + read/write) and
`docx_plus.protection` (document-level lock).

## Building a form: `FormBuilder`

`FormBuilder` is the one place `docx_plus` wraps a document. It seeds an id
registry, materializes the `PlaceholderText` style, and verifies the `w14`
namespace (needed for checkboxes). Add ordinary content through `fb.doc` (a real
python-docx `Document`); add controls through the builder methods.

```python
from docx_plus.controls import FormBuilder

fb = FormBuilder()                 # blank document
# fb = FormBuilder("template.docx")  # open a path
# fb = FormBuilder(existing_doc)     # wrap an open Document

fb.doc.add_heading("New employee form", level=1)   # fb.doc is python-docx
```

Constructor: `FormBuilder(document_or_path=None, *, id_registry=None)`.

Every control is attached **inline to a paragraph you pass in** and returns the
created `<w:sdt>` element (rarely needed). Each takes a `tag=` (machine key,
used later to read/set the value) and an optional `alias=` (the human label
Word shows in the control's title bar).

### Text control

```python
p = fb.doc.add_paragraph("Full name: ")
fb.add_text_control(p, tag="full_name", alias="Full name",
                    placeholder="Type your name")

p = fb.doc.add_paragraph("Notes: ")
fb.add_text_control(p, tag="notes", placeholder="(optional)", multiline=True)
```

`add_text_control(paragraph, *, tag, alias=None, placeholder="Click to enter text", multiline=False)`

### Dropdown and combobox

`items` is either a list of display strings, or a list of `(display, value)`
tuples when the stored value differs from what's shown. A plain dropdown is
closed (must pick from the list); `editable=True` makes it a **combobox** that
also accepts free-form typing.

```python
p = fb.doc.add_paragraph("Department: ")
fb.add_dropdown(p, tag="dept",
                items=[("Engineering", "ENG"), ("Design", "DES"), ("Ops", "OPS")],
                placeholder="Choose a department")

p = fb.doc.add_paragraph("Office: ")
fb.add_dropdown(p, tag="office", items=["New York", "London", "Remote"],
                editable=True)   # combobox: free-form allowed
```

`add_dropdown(paragraph, *, tag, items, alias=None, placeholder="Choose an item", editable=False)`

### Date picker

`date_format` is a Word date pattern (`"M/d/yyyy"`, `"MMMM d, yyyy"`, …).

```python
p = fb.doc.add_paragraph("Start date: ")
fb.add_date_picker(p, tag="start_date", date_format="M/d/yyyy")
```

`add_date_picker(paragraph, *, tag, alias=None, placeholder="Click to select a date", date_format="M/d/yyyy", lcid="en-US")`

### Checkbox

Requires the `w14` namespace on the document root; `FormBuilder` guarantees it
(raises `MissingNamespaceError` otherwise).

```python
p = fb.doc.add_paragraph("Remote-first? ")
fb.add_checkbox(p, tag="remote", checked=False)
```

`add_checkbox(paragraph, *, tag, alias=None, checked=False)`

### Save

```python
fb.save("form.docx")   # returns the path as str
```

## Reading and writing control values

These operate on a plain python-docx `Document` (opened however you like), not
on `FormBuilder`.

```python
from docx import Document
from docx_plus.controls import read_controls, set_control_value, clear_control

doc = Document("form.docx")

# Read everything, keyed by tag (default) — or by="alias"
controls = read_controls(doc)
for tag, ctrl in controls.items():
    print(tag, ctrl.control_type, repr(ctrl.value), ctrl.is_placeholder)

# Set values — type-dispatched on the control:
set_control_value(doc, "full_name", "Ada Lovelace")     # text -> str
set_control_value(doc, "dept", "ENG")                   # dropdown -> stored value OR display text
set_control_value(doc, "remote", True)                  # checkbox -> bool
from datetime import datetime
set_control_value(doc, "start_date", datetime(2026, 6, 1))  # date -> datetime

# Reset one control back to its placeholder state:
clear_control(doc, "notes")

doc.save("form_filled.docx")
```

`read_controls(doc, *, by="tag") -> dict[str, ControlValue]`. A `ControlValue`
is a frozen dataclass with fields `tag`, `alias`, `control_type`, `value`,
`is_placeholder`. `control_type` is one of `"text"`, `"dropdown"`,
`"combobox"`, `"date"`, `"checkbox"`.

`set_control_value(doc, tag, value)` — the accepted `value` type follows the
control type (`str` for text, `str` for dropdown/combobox, `bool` for checkbox,
`datetime` for date). For a closed dropdown, pass either the stored value or the
visible display text; a value matching neither raises `ValueNotInListError`
(comboboxes accept free-form and never raise this).

`clear_control(doc, tag)` — resets to the placeholder state.

## Protecting a form

Lock the document so only the content controls are editable. Idempotent — a
second call replaces the mode.

```python
from docx_plus.protection import protect_document, unprotect_document, is_protected

protect_document(doc, mode="forms")   # only controls editable (the form case)
# other modes: "readOnly", "comments", "trackedChanges"

is_protected(doc)        # -> bool (presence check; does not report which mode)
unprotect_document(doc)  # remove protection (idempotent)
```

`protect_document(doc, *, mode="forms")` — `mode` is one of `"forms"`,
`"readOnly"`, `"comments"`, `"trackedChanges"`. This is **unpassworded**
protection (a UI guard, not encryption).

## End-to-end

```python
from datetime import datetime
from docx import Document
from docx_plus.controls import FormBuilder, read_controls, set_control_value
from docx_plus.protection import protect_document

# --- build ---
fb = FormBuilder()
fb.doc.add_heading("Onboarding", level=1)
p = fb.doc.add_paragraph("Name: ");  fb.add_text_control(p, tag="name")
p = fb.doc.add_paragraph("Team: ");  fb.add_dropdown(p, tag="team",
        items=[("Engineering", "ENG"), ("Design", "DES")])
p = fb.doc.add_paragraph("Start: "); fb.add_date_picker(p, tag="start")
p = fb.doc.add_paragraph("Remote? "); fb.add_checkbox(p, tag="remote")
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

## Errors

All subclass `DocxPlusError` (and the noted builtin).

| Error                     | Raised when                                                        |
| ------------------------- | ------------------------------------------------------------------ |
| `ControlNotFoundError` (`KeyError`)  | `set_control_value` / `clear_control` on an unknown tag |
| `DuplicateTagError` (`ValueError`)   | Two controls share a tag (so `read_controls` by tag is ambiguous) |
| `ValueNotInListError` (`ValueError`) | A closed-dropdown value matches no item (combobox is exempt) |
| `ControlTypeError` (`TypeError`)     | `set_control_value` value type doesn't match the control type |
| `MissingNamespaceError`              | `add_checkbox` on a doc without the `w14` namespace declared |
| `InvalidDropdownItemError` (`TypeError`) | An `items` entry is neither `str` nor `(display, value)` tuple |

See also: `docx_plus/examples/build_form.py` and `populate_form.py`.
