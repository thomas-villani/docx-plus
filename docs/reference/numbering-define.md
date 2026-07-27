# `docx_plus.numbering.define`

Authoring list definitions in `numbering.xml` — the largest remaining
python-docx gap. python-docx has no `CT_AbstractNum` and no `CT_Lvl`
class, so nothing in it can express what a list *looks like*.

OOXML splits a list in two: a `<w:abstractNum>` holds up to nine `<w:lvl>`
children describing each depth, and a `<w:num>` is an *instance* pointing
at one. Paragraphs reference the instance, never the abstract definition
— which is what makes [restarting](numbering-apply.md) possible.

Architecture walkthrough:
[`ARCHITECTURE.md` §7.13](../ARCHITECTURE.md#713-custom-numbering).

!!! tip "Size the hanging indent to the number"

    `hanging` is the width reserved for the number, and the gap between
    number and text is a tab stop at `indent`. If the number is wider
    than `hanging` the tab collapses and a cumulative outline renders
    `1.1.1.On-call lead` rather than `1.1.1. On-call lead`. Deeper levels
    of a `%1.%2.%3.` outline need progressively larger values.

::: docx_plus.numbering.define
    options:
      members:
        - LevelDefinition
        - define_list_definition
        - define_bullet_list
        - define_numbered_list
        - InvalidLevelError
        - MAX_LEVELS
