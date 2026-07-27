# `docx_plus.comments.registry`

Per-document registries for a comment's two id namespaces.

## `CommentIdRegistry` — `w:id`

Comment ids live in a separate uniqueness namespace from SDT, bookmark,
and note ids — comment `5` does not collide with bookmark `5`. The
registry seeds itself from both the comments part and any orphaned
body-side range markers so partially-deleted comments can't trigger id
reuse.

## `DurableIdRegistry` — `w16cid:durableId`

A comment has three identifiers and **only the durable id is stable**:

| Identifier | Where | Stability |
|---|---|---|
| `w:id` | `comments.xml` | A position-dependent index Word renumbers freely |
| `w14:paraId` | comment body paragraph | Changes whenever the body is rewritten |
| `w16cid:durableId` | `commentsIds.xml` | Stable for the life of the comment |

Anything citing a comment from outside the document — a permalink, an
external review tracker, a diff between two revisions — needs the third.
Read it back through
[`AnchoredComment.durable_id`](comments-read.md).

Word writes it as 8 uppercase hex digits (`ST_LongHexNumber`), the same
rendering as `w14:paraId` — verified against a Word-authored file, which
produced values like `33EF1546`. Use
`next_hex()`, not `next()`.

Unlike `paraId`, a durable id is scoped to its one part, so the registry
seeds from `commentsIds.xml` alone.

Architecture walkthrough:
[`ARCHITECTURE.md` §7.6.2](../ARCHITECTURE.md#762-durable-comment-ids-and-author-presence).

::: docx_plus.comments.registry
    options:
      members:
        - CommentIdRegistry
        - DurableIdRegistry
