# `docx_plus.revisions.accept`

Resolve revision marks into final text. Accepting keeps the recorded edit;
rejecting restores the pre-edit state. Run-level insertions and deletions
are handled fully (insertion accept = unwrap, deletion accept = remove, and
the inverses on reject, restoring `<w:delText>` to live `<w:t>`). Move and
property-change marks get a safe, non-structural transform; the one
structural case — a paragraph-mark deletion implies a paragraph merge —
ships a non-corrupting fallback that drops the mark and leaves the text
intact (true merge/split is deferred).

The `*_all` forms process revisions innermost-first so nested marks resolve
before their containers, and are idempotent on a clean document.

::: docx_plus.revisions.accept
    options:
      members:
        - accept_revision
        - reject_revision
        - accept_all_revisions
        - reject_all_revisions
