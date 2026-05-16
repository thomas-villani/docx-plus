# `docx_plus.styles.inspect`

The cascade resolver. Walks the six OOXML formatting layers and returns
a fully-resolved [`ResolvedFormatting`][docx_plus.styles.inspect.ResolvedFormatting]
plus optional per-field provenance.

See [`ARCHITECTURE.md` §2](../ARCHITECTURE.md#2-the-cascade-resolver) for
the algorithm walkthrough and the toggle semantics.

::: docx_plus.styles.inspect
    options:
      members:
        - resolve_effective_formatting
        - ResolvedFormatting
        - FormattingSource
        - StyleCascadeError
        - MissingPartError
