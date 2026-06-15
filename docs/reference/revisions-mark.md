# `docx_plus.revisions.mark`

Author tracked insertions and deletions by wrapping run(s) already in the
document. `mark_insertion` wraps the target in `<w:ins>`; `mark_deletion`
wraps it in `<w:del>` and retags each `<w:t>` to `<w:delText>`. Both take
the same target shapes as `comments.add_comment` — a single run, a whole
paragraph, or a `(start_run, end_run)` range — but a range must lie within
one paragraph, since `w:ins` / `w:del` cannot span a paragraph boundary.

Runs wrapped in a revision are not visible through python-docx's
`paragraph.runs`; read them back with `read_revisions`.

::: docx_plus.revisions.mark
    options:
      members:
        - mark_insertion
        - mark_deletion
        - RevisionRef
        - RevisionTarget
        - RevisionNotFoundError
