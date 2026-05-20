# `docx_plus.publishing.toc`

Table of Contents via the `TOC` complex field. The instruction
string is assembled from the `levels`, `hyperlink`, and
`page_numbers` kwargs to match what Word's "Insert → Table of
Contents" UI produces. Word populates the visible body of the TOC
on next open; call
[`docx_plus.fields.mark_fields_dirty`](fields-update.md) before
saving so the recalculation actually fires.

Architecture walkthrough: [`ARCHITECTURE.md` §7.10](../ARCHITECTURE.md#710-publishing).

::: docx_plus.publishing.toc
    options:
      members:
        - add_toc
