# `docx_plus.styles.sweep`

The document-wide cascade sweep. Resolves every paragraph and run against a
single shared cache, in document order.

[`resolve_effective_formatting`][docx_plus.styles.inspect.resolve_effective_formatting]
answers "what does *this* paragraph render as", rebuilding the theme, the
styles part, and each `basedOn` chain on every call.
[`iter_resolved_paragraphs`][docx_plus.styles.sweep.iter_resolved_paragraphs]
answers the same question about every paragraph at once, resolving those
document-level inputs once for the whole walk — roughly 5x faster per target
on a text-heavy document, and the read half any whole-document analysis
needs.

Pass `include_baseline=True` to also resolve each target with its own direct
formatting excluded, populating `.baseline`. That is the comparison behind
"is this direct property doing anything?" — see
[`stop_below`](styles-inspect.md) for what the baseline is a resolve of.

::: docx_plus.styles.sweep
    options:
      members:
        - iter_resolved_paragraphs
        - ResolvedParagraph
        - ResolvedRun
