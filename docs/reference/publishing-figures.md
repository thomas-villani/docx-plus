# `docx_plus.publishing.figures`

Table of Figures via the `TOC \c "<caption_type>"` complex field.
Structurally identical to a Table of Contents but driven by `SEQ`
caption names instead of paragraph outline levels — the
`caption_type` must match the value passed to
[`docx_plus.publishing.add_caption`](publishing-captions.md).

Architecture walkthrough: [`ARCHITECTURE.md` §7.10](../ARCHITECTURE.md#710-publishing).

::: docx_plus.publishing.figures
    options:
      members:
        - add_table_of_figures
