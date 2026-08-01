# `docx_plus.controls.builder`

Build content controls (SDTs) — text, dropdown, date, checkbox.
`FormBuilder` wraps a python-docx `Document` and emits valid `w:sdt`
blocks that round-trip through Word.

Architecture walkthrough: [Content controls](../concepts/controls.md).

::: docx_plus.controls.builder
    options:
      members:
        - FormBuilder
        - DropdownItem
        - MissingNamespaceError
        - InvalidDropdownItemError
