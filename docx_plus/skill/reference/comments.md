# Comments — anchored, threaded review comments

Module: `docx_plus.comments`. Add, read, edit, and delete Word review comments
that are **anchored** to specific text, plus reply to them and resolve or reopen
the thread.

**Why this exists:** python-docx's `Comments.add_comment` writes only the
part-side comment body, not the body-side range markers. The result is a comment
that exists but isn't tied to any text, so Word's "show in document" can't jump
to it. `docx_plus.add_comment` writes all of it — the `commentRangeStart` /
`commentRangeEnd` markers *and* the `commentReference` run *and* the
`<w:comment>` body — so the comment actually highlights its target.

## Adding comments

The target can be a single `Run`, a whole `Paragraph` (every run is wrapped;
needs ≥1 run), or a `(start_run, end_run)` tuple for a multi-run range.

```python
from docx_plus.comments import add_comment

doc = Document()
p = doc.add_paragraph()
p.add_run("Project Apollo ")
target = p.add_run("ships next quarter")
p.add_run(".")

ref = add_comment(target, "Optimistic — let's see what QA says.",
                  author="Alice", initials="A")
print(ref.comment_id)
```

`add_comment(target, text, *, author="", initials=None, id_registry=None,
para_id_registry=None) -> CommentRef` (`CommentRef` has `comment_id`,
`body_element`). Every comment is written as an unresolved thread root, so it
can be replied to and resolved immediately.

### Adding several at once — share the registry

When adding multiple comments in one session, build one `CommentIdRegistry` and
pass it to every call so the allocated ids stay unique:

```python
from docx_plus.comments import CommentIdRegistry, add_comment

reg = CommentIdRegistry(doc)
add_comment(run_a, "First.",  author="Alice", initials="A", id_registry=reg)
add_comment(para_b, "Second.", author="Bob",   initials="B", id_registry=reg)
add_comment((start, end), "Range.", author="Carol", initials="C", id_registry=reg)
```

## Threads — replies, resolve, reopen

Word has modelled comments as threads since 2013: a root plus replies, with a
resolved flag. The thread graph lives in a second part
(`commentsExtended.xml`) that python-docx does not touch at all.

```python
from docx_plus.comments import reply_to_comment, resolve_comment, reopen_comment

root = add_comment(target, "Where does this number come from?", author="Alice")
reply_to_comment(doc, root.comment_id, "The Q2 capacity model.", author="Bob")
reply_to_comment(doc, root.comment_id, "Thanks.", author="Alice")

resolve_comment(doc, root.comment_id)   # closes the thread in Word's review pane
reopen_comment(doc, root.comment_id)    # exact inverse
```

- `reply_to_comment(doc, parent_id, text, *, author="", initials=None,
  id_registry=None, para_id_registry=None) -> CommentRef` — the reply spans the
  same text range as its parent, which is how Word renders a thread as one
  balloon. Raises `CommentNotFoundError` if `parent_id` is unknown.
- `resolve_comment(doc, comment_id)` / `reopen_comment(doc, comment_id)` —
  resolution is **thread-wide**, matching Word's Resolve button, so naming any
  member moves the whole thread.

Documents from python-docx or pre-2013 Word have no thread data; they read as
one unresolved single-comment thread each, and replying to or resolving one
upgrades it in place.

## Reading comments

```python
from docx_plus.comments import read_comments, read_threads

for c in read_comments(doc):
    print(f"[{c.author}] {c.text!r} on {c.anchored_text!r} (p{c.paragraph_index})")

for t in read_threads(doc):
    state = "resolved" if t.resolved else "open"
    print(f"[{state}] {t.root.author}: {t.root.text}")
    for reply in t.replies:
        print(f"    -> {reply.author}: {reply.text}")
```

`read_comments(doc) -> list[AnchoredComment]`. An `AnchoredComment` is a frozen
dataclass with `comment_id`, `author`, `initials`, `timestamp`, `text`,
`anchored_text` (the document text the comment is anchored to),
`paragraph_index`, `parent_id` (`None` for a thread root), and `resolved`.

`read_threads(doc) -> list[CommentThread]` returns the same comments grouped —
one `CommentThread` per root, with `root`, `replies`, and `resolved`.

## Editing and deleting

```python
from docx_plus.comments import edit_comment, delete_comment, clear_all_comments

edit_comment(doc, ref.comment_id, "Revised note.")  # body text only; keeps
                                                     # author/date/initials/anchors
delete_comment(doc, ref.comment_id)   # removes the comment and its replies
clear_all_comments(doc)               # delete every comment; idempotent
```

- `edit_comment(doc, comment_id, text)` — replaces body text in place; preserves
  `w:author` / `w:date` / `w:initials`, the body anchors, and the comment's
  place in its thread. Raises `CommentNotFoundError` (`KeyError`) on an unknown
  id.
- `delete_comment(doc, comment_id, *, include_replies=True)` — removes the range
  markers, the reference run, the body, and the thread entry. By default it also
  deletes the reply subtree, as Word does; `include_replies=False` promotes the
  orphaned replies to roots. Idempotent (unknown id is a no-op).
- `clear_all_comments(doc, *, remove_part=False)` — scrubs every comment and
  thread entry. `remove_part=True` also tears down the underlying parts.

## End-to-end

```python
from docx import Document
from docx_plus.comments import add_comment, read_comments

doc = Document()
p = doc.add_paragraph()
p.add_run("The migration ")
target = p.add_run("completes in Q3")
p.add_run(".")

add_comment(target, "Confirm with the platform team.", author="Reviewer")
doc.save("review.docx")

for c in read_comments(Document("review.docx")):
    print(f"{c.author}: {c.text!r} -> {c.anchored_text!r}")
```

Type alias: `CommentTarget = Run | Paragraph | tuple[Run, Run]`.

## Citing a comment across edits — `durable_id`

A comment has three identifiers and **only one is stable**:

| Identifier | Stability |
|---|---|
| `comment_id` (`w:id`) | A position-dependent index Word renumbers freely |
| `w14:paraId` | Changes whenever the comment body is rewritten |
| `durable_id` (`w16cid:durableId`) | Stable for the life of the comment |

If you are storing a reference to a comment anywhere outside the
document — a review tracker, a permalink, a diff between two revisions
— use `durable_id`. Anything else will silently point at the wrong
comment later.

```python
for c in read_comments(doc):
    print(c.durable_id, c.text)      # e.g. '33EF1546'
```

Written automatically by `add_comment` / `reply_to_comment`. `None` on
documents written before v0.5, by python-docx, or by anything other
than Word 2016+. Share a `DurableIdRegistry` via
`durable_id_registry=` when inserting many comments at once.

## Author presence — `people.xml`

Optional and **cosmetic**: it drives the presence dot beside a comment
in Word's reviewing pane. Comments, threading, and resolution all work
without it, so `add_comment` deliberately does not write it — doing so
would mean inventing a directory identity for an author the library
knows nothing about.

```python
from docx_plus.comments import set_author_presence, read_author_presence

set_author_presence(doc, "Reviewer")                    # providerId="None"
set_author_presence(doc, "Ana Silva", provider_id="AD",
                    user_id="S::ana@example.com::4f2c...")

read_author_presence(doc)   # -> [AuthorPresence(author=..., provider_id=..., user_id=...)]
```

The **author name is the only join** to `comments.xml` — pass exactly
the string used as `author=` on that person's comments.

Stale authors are not pruned when their comments are deleted (Word
doesn't either). `clear_author_presence(doc, remove_part=True)` is the
escape hatch.

From the shell: `docx-plus comments list FILE [--unresolved] [--json]`, and
`docx-plus comments resolve|reopen FILE ID -o OUT`. See `reference/cli.md`.

See also: `docx_plus/examples/add_comments.py`,
`docx_plus/examples/threaded_comments.py`.
