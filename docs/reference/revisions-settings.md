# `docx_plus.revisions.settings`

Toggle document-wide track-changes mode by writing or removing
`<w:trackChanges/>` in `settings.xml`. When present, Word records every
edit the reader makes as a revision. This is independent of authoring
specific revision marks with `mark_insertion` / `mark_deletion`, which
write marks regardless of the flag.

Both functions are idempotent: enabling collapses to a single element,
disabling removes every copy.

::: docx_plus.revisions.settings
    options:
      members:
        - enable_track_changes
        - disable_track_changes
