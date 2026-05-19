# `docx_plus.notes.registry`

Footnote and endnote id registries. The two kinds use disjoint
namespaces — a footnote with id `1` and an endnote with id `1` can
coexist. Ids `-1` (separator) and `0` (continuation separator) are
reserved by Word; the registries refuse to issue or reserve them
(the underlying range check on `_IdRegistryBase.reserve` rejects values
outside `[1, 2**31 - 1]` before any duplicate check runs).

::: docx_plus.notes.registry
    options:
      members:
        - FootnoteIdRegistry
        - EndnoteIdRegistry
