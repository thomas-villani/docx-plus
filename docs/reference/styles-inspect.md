# `docx_plus.styles.inspect`

The cascade resolver. Walks the OOXML formatting layers and returns a
fully-resolved [`ResolvedFormatting`][docx_plus.styles.inspect.ResolvedFormatting]
plus optional per-field provenance.

See [`ARCHITECTURE.md` §2](../ARCHITECTURE.md#2-the-cascade-resolver) for
the algorithm walkthrough and the toggle semantics.

## Toggle properties do not override

`bold`, `italic` and the ten other ECMA-376 17.7.3 toggles combine by a
rule that catches most people out, so it is worth stating up front:

- A `basedOn` chain is **one level** and flattens by plain override. A
  child style restating its parent's `<w:b/>` stays bold.
- A paragraph style and a character style are **two levels**, and two
  levels both asking for bold cancel. This is the spec's own example.
- `docDefaults` is the **base**, not a level, and a style restating the
  base is inert.
- Direct formatting on a run is **absolute** — it states the value rather
  than flipping it.

The rule was settled by measuring live Word, because the spec prose admits
several incompatible readings and this library previously shipped one of
the wrong ones. `tests/test_cascade_word_verified.py` holds the
measurements.

## `include_provenance` vs `stop_below`

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
