# `docx_plus.styles.inspect`

The cascade resolver. Walks the six OOXML formatting layers and returns
a fully-resolved [`ResolvedFormatting`][docx_plus.styles.inspect.ResolvedFormatting]
plus optional per-field provenance.

See [`ARCHITECTURE.md` §2](../ARCHITECTURE.md#2-the-cascade-resolver) for
the algorithm walkthrough and the toggle semantics.

Two features answer different questions and are easy to confuse:

- **`include_provenance`** names the layer each value *came from*.
- **`stop_below`** re-resolves with a layer excluded, giving the value
  that *would have surfaced without it*.

Provenance alone cannot answer the second: knowing a run's size came from
`directRun` does not say what deleting that `<w:rPr>` would leave. Together
they are what lets a caller tell a direct property that changes nothing
from one that genuinely overrides its style — the basis of every
consistency rule in [`docx_plus.lint`](lint.md).

::: docx_plus.styles.inspect
    options:
      members:
        - resolve_effective_formatting
        - ResolvedFormatting
        - TableContext
        - FormattingSource
        - StyleCascadeError
        - MissingPartError
