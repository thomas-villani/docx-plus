# `docx_plus.numbering.apply`

Attaching list definitions to paragraphs. A paragraph joins a list by
carrying `<w:numPr>` in its `<w:pPr>`, naming a `w:numId` and a `w:ilvl`.
Paragraphs sharing a `numId` continue one sequence, in document order.

Restarting deserves a note: it is **not** a paragraph property in OOXML.
There is nowhere to say "count from 1 again here". `restart_list` does
what Word does — adds a second `<w:num>` over the same `<w:abstractNum>`
carrying a `<w:startOverride>`, giving an independent counter that looks
identical.

Architecture walkthrough:
[Custom numbering](../concepts/numbering.md).

::: docx_plus.numbering.apply
    options:
      members:
        - apply_list
        - remove_list
        - restart_list
        - ListDefinitionNotFoundError
