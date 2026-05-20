# `docx_plus.publishing.captions`

Figure / table captions: a literal label run (`"Figure "`,
`"Table "`, etc.) followed by a `SEQ` complex field that
auto-numbers items sharing the same caption type. The caption type
is the same name a downstream Table of Figures uses to find
captions (see
[`docx_plus.publishing.figures`](publishing-figures.md)).

Architecture walkthrough: [`ARCHITECTURE.md` §7.10](../ARCHITECTURE.md#710-publishing).

::: docx_plus.publishing.captions
    options:
      members:
        - add_caption
