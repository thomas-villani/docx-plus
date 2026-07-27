# `docx_plus.comments.people`

Comment author presence — `/word/people.xml`.

Word records one entry per comment author, driving the presence
indicator beside a comment in the reviewing pane and the identity a
click resolves to:

```xml
<w15:people>
  <w15:person w15:author="Thomas Villani">
    <w15:presenceInfo w15:providerId="AD"
                      w15:userId="S::thomas@example.com::541bd2ef-..."/>
  </w15:person>
</w15:people>
```

!!! note "Nothing here runs automatically"
    The part is **purely cosmetic** — comments, threading, and
    resolution all work without it, and Word neither requires it nor
    complains when it is missing.

    [`add_comment`](comments-anchor.md) deliberately does *not* write
    it. Registering an author means inventing a `userId` for someone
    the library knows nothing about, and a fabricated directory
    identity is worse than an absent one. Call `set_author_presence`
    explicitly when you want the entry.

The **author name is the only join** between this part and
`comments.xml` — there are no comment ids here — so the name passed to
`set_author_presence` must match the `w:author` on that person's
comments exactly.

Stale authors are **not** pruned when their last comment is deleted.
Word does not prune them either, and doing so would need the author
ref-counted across every surviving comment. `clear_author_presence` is
the explicit escape hatch.

Architecture walkthrough:
[`ARCHITECTURE.md` §7.6.2](../ARCHITECTURE.md#762-durable-comment-ids-and-author-presence).

::: docx_plus.comments.people
    options:
      members:
        - AuthorPresence
        - set_author_presence
        - read_author_presence
        - clear_author_presence
        - LOCAL_PROVIDER
