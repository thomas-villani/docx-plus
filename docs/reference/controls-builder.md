# `docx_plus.controls.builder`

Build content controls (SDTs) — text, dropdown, date, checkbox.
`FormBuilder` wraps a python-docx `Document` and emits valid `w:sdt`
blocks that round-trip through Word.

Architecture walkthrough: [`ARCHITECTURE.md` §6](../ARCHITECTURE.md#6-content-controls).

::: docx_plus.controls.builder
    options:
      members:
        - FormBuilder
        - DropdownItem
        - MissingNamespaceError
        - InvalidDropdownItemError
