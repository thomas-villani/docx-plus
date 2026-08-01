# Anchored comments

`comments/anchor.py:add_comment(target, text, ...)` is the v0.2
headline. Closes the largest python-docx gap: python-docx 1.x writes
`<w:comment>` into `comments.xml` but skips the three body-side
elements that anchor the comment to a text range, so its comments show
in the review pane but have nothing to point at when the reader clicks
"show in document".

For the calls, see the [comments guide](../guides/comments.md).

## The five elements per comment

Each `add_comment` writes five elements:

1. `<w:commentRangeStart w:id=N/>` — placed before `start_anchor` via
   `addprevious`
2. `<w:commentRangeEnd w:id=N/>` — placed after `end_anchor` via
   `addnext`
3. The reference run — `<w:r><w:rPr><w:rStyle val="CommentReference"/></w:rPr><w:commentReference w:id=N/></w:r>`
   — placed after the range end
4. The comment body — `<w:comment w:id=N w:author=... w:date=... [w:initials=...]>`
   appended to the root of `comments.xml` (via `get_or_create_part`), its
   paragraphs stamped with `w14:paraId`
5. The thread entry — `<w15:commentEx w15:paraId=P w15:done="0"/>`
   appended to `commentsExtended.xml` (v0.4; see below)

Target shapes: a python-docx `Run` (brackets just that run), a
`Paragraph` (brackets every run, must have ≥1 run), or a
`(start_run, end_run)` tuple for a range. Range tuples may span
paragraphs; OOXML permits this. Comment body uses
`xml:space="preserve"` so leading/trailing whitespace survives Word's
XML reader.

`delete_comment(doc, comment_id, *, include_replies=True)` is the
inverse — removes all five elements and is idempotent (missing id is a
no-op). By default it also deletes the comment's reply subtree, which is
what Word does when you delete a thread root.

`read_comments(doc)` walks `comments.xml` and pairs each `<w:comment>`
with its body range, extracting `author`, `initials`, `timestamp`
(parsed `xsd:dateTime`), the comment `text`, the `anchored_text`
between the body markers, and the `paragraph_index` where the
`commentRangeStart` sits. Orphaned comments (no matching body range)
appear with `anchored_text=""` and `paragraph_index=-1`.

`CommentIdRegistry` lives in its own namespace (separate from SDT,
bookmark, note ids). It seeds from both the comments part AND any
orphaned body-side anchors so a partially-deleted comment cannot
trigger id reuse.

## Threading — `commentsExtended.xml`

Word 2013 made comments *threaded* without touching `<w:comment>`: the
thread graph went into a second part, `/word/commentsExtended.xml`,
holding one `<w15:commentEx>` per comment with an optional
`w15:paraIdParent` and a `w15:done` resolved flag. `comments/threads.py`
(v0.4) is the public surface — `reply_to_comment`, `resolve_comment`,
`reopen_comment`, `read_threads` — over `comments/_extended.py`, which
owns the part.

Three properties of Microsoft's design shape the implementation:

1. **Entries key off `w14:paraId`, not `w:id`.** The key is the `paraId`
   of the comment body's *last* paragraph. Comment ids and thread keys
   are separate namespaces, so every mapping between them routes through
   the comment body. `ParaIdRegistry` (`core/ids.py`) allocates the
   values; unlike every other registry it is unique across the whole
   *package*, not one part, so it seeds from the body plus the comments /
   footnotes / endnotes parts. Word's legal range for a `paraId` —
   nonzero and below `0x80000000` — is exactly the existing 31-bit
   allocator range, so only the hex rendering is new.
2. **A reply shares its parent's anchor range, and marker order is
   display order.** `_mirror_anchors` nests the markers the way Word
   does — every member's `commentRangeStart` before the text, each
   `commentRangeEnd` + reference-run pair after it. Word sorts a thread's
   balloons by where each *reference mark* sits in the body, not by date
   or by position in `comments.xml`, so a new reply's markers append
   after every marker the thread already owns. Inserting them beside the
   parent's pair instead renders each thread in reverse chronological
   order — a defect caught only by opening the output in Word, and pinned
   now by `test_replies_are_appended_in_conversation_order`. A parent
   with no anchors (an orphaned comment) leaves the reply orphaned too
   rather than inventing a range.
3. **Resolution is thread-wide.** Word's Resolve button greys out root
   and replies together, so `resolve_comment` sets `w15:done` across the
   whole thread no matter which member you name.

The part is optional in the format, and every reader here treats its
absence as "one unresolved root per comment" — which is the correct
reading of a document from python-docx or pre-2013 Word. The write paths
materialize the missing metadata in place, so replying to or resolving a
foreign comment upgrades it rather than failing.

`core/parts.py` supplies `COMMENTS_EXTENDED_SPEC` plus the
`CT_COMMENTS_EXTENDED` / `RT_COMMENTS_EXTENDED` URIs — Microsoft
extensions with no member in python-docx's `CT` / `RT` enums — and
registers an `XmlPart` subclass for the content type so an existing
extended part deserializes with a parsed `.element` instead of a blob.

Because `w14:paraId` is now written into `comments.xml`, the fabricated
comments root declares `xmlns:w14` and `mc:Ignorable="w14"`. And because
`w15` belongs only to the extended part, `core/ns.py` splits the
namespace map in two: `NSMAP` is the *query* map that XPath binds, while
the narrower `BUILD_NSMAP` is what `el()` declares on main-document
elements. An element outside those prefixes declares just its own, so
adding `w15` did not stamp a stray `xmlns:w15` onto every element the
library writes.

## Durable ids and author presence

v0.5 added the last two comment side-parts Word writes. Every URI and
element shape below was verified against a file **Word 2016 authored
itself** — driven over COM, saved, unzipped, and read — not inferred
from the spec.

### `commentsIds.xml` — the only stable identifier

A comment has three ids and only one of them survives an edit:

| Identifier | Where | Stability |
|---|---|---|
| `w:id` | `comments.xml` | A position-dependent index Word renumbers |
| `w14:paraId` | body paragraph | Changes when the body is rewritten |
| `w16cid:durableId` | `commentsIds.xml` | Stable for the comment's life |

Anything citing a comment from outside the document — a permalink, an
external review tracker, a diff between two revisions — needs the third,
which is why Word 2016 added the part.

Two facts about it were **wrong in the original plan** and are worth
recording, since both would have shipped:

1. **`durableId` is hex, not decimal.** It is `ST_LongHexNumber` — the
   same 8-uppercase-digit rendering as `w14:paraId`. Word wrote
   `33EF1546` / `31436C50` / `50E18CF9`. The plan called for a decimal
   collector and a decimal registry; instead `DurableIdRegistry` reuses
   the existing `_collect_hex_id_attrs` / `next_hex` machinery, so the
   feature added *less* core code than budgeted, not more.
2. **There is a fifth part.** `commentsExtensible.xml` (`w16cex`, 2018)
   keys off `durableId` and carries `dateUtc`, because `w:comment/@w:date`
   is local time. Out of scope here: `commentsIds` predates it by two
   years, so writing one without the other is a state Word itself
   produced for years — and Word did not add it when resaving a file of
   ours that lacked it. Tracked in `ROADMAP.md`.

Entries key off `w14:paraId` exactly as `commentsExtended.xml` does, so
`comments/_ids.py` reuses `_extended.py`'s `stamp_para_ids` /
`thread_key` / `key_maps` rather than building a second bridge from
comment ids to part entries. `upsert_comment_id` never reissues an
existing id — that would defeat the part's whole purpose, breaking every
reference already taken against the old value.

Writing is automatic (`add_comment`, `reply_to_comment`) because Word
regenerates missing entries anyway, so emitting them moves output toward
native rather than away from it.

### `people.xml` — presence, and why it is opt-in

`<w15:person w15:author="…">` carries a `<w15:presenceInfo>` child with
a `providerId` (`"AD"`, `"Windows Live"`, `"Office365"`, or `"None"`)
and a provider-scoped `userId`. It drives the presence dot beside a
comment in the reviewing pane.

**`add_comment` deliberately does not write it.** Registering an author
means inventing a `userId` for someone the library knows nothing about,
and a fabricated directory identity is worse than an absent one. The
part is purely cosmetic — comments, threading, and resolution all work
without it — so `comments/people.py` exposes it as an explicit
`set_author_presence` call instead.

The **author name is the only join** to `comments.xml`; the part carries
no comment ids. And stale authors are **not** pruned on delete: Word
does not prune them either, and doing so would need the author
ref-counted across every surviving comment.

### Verified round-trip

Opening a document this library wrote, then resaving it from Word,
preserved both `paraId` and `durableId` byte-for-byte, and preserved
`people.xml` including a non-default `providerId="AD"` entry. Word added
no parts of its own.
