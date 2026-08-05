# `docx_plus.controls.read`

Read and modify the values of existing content controls.
`list_controls(doc)` returns every control as a `list[ControlValue]` in
document order; `read_controls(doc)` is the keyed convenience on top of it,
returning a `dict[str, ControlValue]` keyed by tag (default) or alias.
`set_control_value` / `clear_control` mutate single controls.

!!! warning "Reach for `list_controls` on Word-authored documents"

    `w:tag` is optional and non-unique in OOXML, and Word writes
    `<w:tag w:val=""/>` for any control the author did not explicitly tag.
    `read_controls` can only report controls that have a usable key, so on a
    typical Word form it omits most of them. `list_controls` reports every
    control, with `control_id` for identity.

The five typed errors are all dual-base (`DocxPlusError` plus a stdlib
exception) so callers can match either contract — see
[the error hierarchy](../concepts/invariants.md#error-hierarchy).

::: docx_plus.controls.read
    options:
      members:
        - ControlValue
        - ControlType
        - WRITABLE_TYPES
        - list_controls
        - read_controls
        - set_control_value
        - clear_control
        - ControlNotFoundError
        - DuplicateTagError
        - ValueNotInListError
        - ControlTypeError
