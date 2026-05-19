# `docx_plus.fields.update`

Mark every field in a document for recalculation on the next Word
open. Sets `<w:updateFields w:val="true"/>` in `settings.xml`. Word
flips the flag back to `false` after recalculating — it's a one-shot
mechanism, not persistent state.

`mark_fields_dirty` is idempotent: a second call updates the existing
element rather than duplicating it.

::: docx_plus.fields.update
    options:
      members:
        - mark_fields_dirty
