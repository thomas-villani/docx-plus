# `docx_plus.revisions.registry`

Per-document registry of issued revision `w:id` values. Unlike comments,
bookmarks, and notes — each of which has its own id namespace — *all*
tracked-change element types (`w:ins`, `w:del`, the move wrappers and
their range markers, `w:rPrChange`, `w:pPrChange`) share a **single**
`w:id` namespace, so a `w:ins` and a `w:del` cannot reuse the same id.
The registry seeds itself from every revision-bearing element in the body.

::: docx_plus.revisions.registry
    options:
      members:
        - RevisionIdRegistry
