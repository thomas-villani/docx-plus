# Tracked changes

`docx_plus.revisions` authors, reads, and resolves Word tracked changes —
revision marks python-docx cannot touch at all. It can write runs but
cannot mark them as tracked insertions or deletions, read existing
revisions, or accept and reject them.

Unlike comments, these elements (`w:ins`, `w:del` with `w:delText`, the
move wrappers, and the property-change markers) live **inline in the
document body**, so there is no separate part involved.

!!! warning "Wrapped runs vanish from `paragraph.runs`"
    python-docx's `paragraph.runs` only sees direct `w:r` children. After
    `mark_insertion(run)`, that run is inside a `<w:ins>` and no longer
    appears in the list. Read it back with `read_revisions` instead — and
    note it reappears once `accept` or `reject` unwraps it.

## Turning track-changes mode on

This is the document-wide `<w:trackChanges/>` flag: it tells Word to record
edits the *reader* makes. It is independent of the authoring marks below,
which work whether or not it is set.

```python
from docx_plus.revisions import disable_track_changes, enable_track_changes

enable_track_changes(doc)     # idempotent
disable_track_changes(doc)    # idempotent; leaves existing marks alone
```

## Marking insertions and deletions

Wrap run(s) that are **already in the document**. Target shapes match
`add_comment`: a single `Run`, a whole `Paragraph` (needs ≥1 run), or a
`(start_run, end_run)` tuple.

```python
from docx_plus.revisions import RevisionIdRegistry, mark_deletion, mark_insertion

p = doc.add_paragraph()
p.add_run("The plan ")
new = p.add_run("ships in Q3 ")
old = p.add_run("ships someday")

reg = RevisionIdRegistry(doc)                          # share across the batch
mark_insertion(new, author="Alice", id_registry=reg)
mark_deletion(old, author="Bob", id_registry=reg)      # retags w:t -> w:delText
```

`mark_insertion(target, *, author="", date=None, id_registry=None)` returns
a `RevisionRef` with `revision_id` and `body_element`; `mark_deletion` has
the same signature. `date` defaults to now in UTC at millisecond precision.

!!! note "A range cannot cross a paragraph boundary"
    `w:ins` and `w:del` are inline elements, so a `(start_run, end_run)`
    tuple must stay inside one paragraph. (Comment ranges *can* span
    paragraphs — the two differ here.)

All revision types share **one** id namespace, so a `w:ins` id and a
`w:del` id cannot collide. Use one `RevisionIdRegistry` per editing
session.

## Reading revisions

```python
from docx_plus.revisions import read_revisions

for rv in read_revisions(doc):
    print(rv.revision_type, rv.author, repr(rv.text), rv.paragraph_index)
```

`read_revisions(doc)` returns a `list[TrackedChange]` in document order.
`TrackedChange` is a frozen dataclass with `revision_id`, `revision_type`,
`author`, `timestamp`, `text`, and `paragraph_index`.

`revision_type` is one of `insertion`, `deletion`, `move_from`, `move_to`,
`format_run`, `format_paragraph`, `paragraph_mark_insertion`,
`paragraph_mark_deletion`. Insertion text comes from `<w:t>` and deletion
text from `<w:delText>`; format and paragraph-mark changes carry empty
`text`.

## Accepting and rejecting

```python
from docx_plus.revisions import (
    accept_all_revisions,
    accept_revision,
    reject_all_revisions,
    reject_revision,
)

accept_revision(doc, rev_id)     # keep the edit
reject_revision(doc, other_id)   # restore the pre-edit state

accept_all_revisions(doc)        # idempotent
reject_all_revisions(doc)        # ...so this then finds nothing left
```

| Revision | Accept | Reject |
|---|---|---|
| Insertion | Keep the text (unwrap) | Drop it |
| Deletion | Remove the text | Restore it as live `<w:t>` |
| Move / format change | Safe mechanical transform | Inverse |
| Paragraph-mark | Non-corrupting fallback: the mark is dropped, text left intact | Same |

True paragraph merge and split on a paragraph-mark revision is deferred.

Each call **consumes** the revision, so accept and reject are alternatives
for one id, not a sequence — calling both raises `RevisionNotFoundError`
(which is also a `KeyError`). The `*_all` forms process innermost marks
first, so nested revisions resolve cleanly.

## End to end

```python
from docx import Document
from docx_plus.revisions import (
    accept_all_revisions,
    enable_track_changes,
    mark_deletion,
    mark_insertion,
    read_revisions,
)

doc = Document()
enable_track_changes(doc)
p = doc.add_paragraph()
p.add_run("Keep ")
mark_insertion(p.add_run("this new bit "), author="A")
mark_deletion(p.add_run("this old bit"), author="B")
doc.save("draft.docx")

reopened = Document("draft.docx")
print([(r.revision_type, r.text) for r in read_revisions(reopened)])

accept_all_revisions(reopened)
print(reopened.paragraphs[0].text)     # -> "Keep this new bit "
```

Type alias: `RevisionTarget = Run | Paragraph | tuple[Run, Run]`.

## See also

- [How tracked changes work](../concepts/revisions.md)
- Reference: [`revisions.mark`](../reference/revisions-mark.md),
  [`revisions.read`](../reference/revisions-read.md),
  [`revisions.accept`](../reference/revisions-accept.md),
  [`revisions.settings`](../reference/revisions-settings.md)
- Example: `track_changes.py`
