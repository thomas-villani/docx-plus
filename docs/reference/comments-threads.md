# `docx_plus.comments.threads`

Threaded comments — the model Word has used since 2013 and python-docx
does not expose. A thread is a root comment plus its replies, carrying a
resolved flag that drives the review pane's **Resolve** button.

The thread graph lives in a second OOXML part,
`/word/commentsExtended.xml`, whose `<w15:commentEx>` entries key off the
`w14:paraId` of each comment body's **last paragraph** — not off the
comment's `w:id`. `add_comment` stamps that `paraId` and writes an
unresolved entry for every comment it inserts, so any comment is
immediately reply-able and resolvable.

Resolution is thread-wide: naming any comment in a thread resolves or
reopens the whole thread, matching Word's UI.

A document with no extended part — anything written by python-docx, or by
Word before 2013 — reads as one unresolved single-comment thread per
comment. Replying to or resolving such a comment materializes the missing
metadata in place.

::: docx_plus.comments.threads
    options:
      members:
        - reply_to_comment
        - resolve_comment
        - reopen_comment
        - read_threads
        - CommentThread
