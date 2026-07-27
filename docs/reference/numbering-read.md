# `docx_plus.numbering.read`

Reading list definitions back out of `numbering.xml`, including ones Word
or another tool wrote.

Like every reader in the library this never fabricates a part: a document
with no `numbering.xml` reads as an empty list rather than gaining one as
a side effect of being inspected.

!!! note "A fresh `Document()` is not empty"

    python-docx's bundled template ships nine `abstractNum` entries and
    nine `num` instances backing the built-in `List Bullet` and
    `List Number` styles, so an untouched document already reports nine
    definitions.

::: docx_plus.numbering.read
    options:
      members:
        - read_list_definitions
        - ListDefinition
        - ListLevel
