# `docx_plus.controls.read`

Read and modify the values of existing content controls.
`read_controls(doc)` returns a flat `dict[str, ControlValue]` keyed by
tag (default) or alias; `set_control_value` / `clear_control` mutate
single controls.

The five typed errors are all dual-base (`DocxPlusError` plus a stdlib
exception) so callers can match either contract — see
[`ARCHITECTURE.md` §9](../ARCHITECTURE.md#9-error-hierarchy).

::: docx_plus.controls.read
    options:
      members:
        - ControlValue
        - ControlType
        - read_controls
        - set_control_value
        - clear_control
        - ControlNotFoundError
        - DuplicateTagError
        - ValueNotInListError
        - ControlTypeError
