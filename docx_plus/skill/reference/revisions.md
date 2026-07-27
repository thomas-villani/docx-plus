# Revisions — tracked changes (insertions, deletions, accept/reject)

Module: `docx_plus.revisions`. Author, read, and resolve Word tracked changes —
the revision marks python-docx cannot touch at all.

**Why this exists:** python-docx can write runs but cannot mark them as tracked
insertions or deletions, read existing revisions, or accept/reject them. These
elements (`w:ins`, `w:del` with `w:delText`, move wrappers, property-change
markers) live **inline in the document body** — there is no separate part (unlike
comments). `docx_plus.revisions` writes and resolves them correctly.

**Key gotcha:** runs wrapped in `w:ins` / `w:del` are **invisible** to
python-docx's `paragraph.runs` (it only sees direct `w:r` children). After
`mark_insertion(run)`, `paragraph.runs` no longer lists that run — read it back
with `read_revisions`. After `accept`/`reject` unwraps it, it reappears.

## Turning track-changes mode on/off

This is the document-wide `<w:trackChanges/>` flag in `settings.xml` — it tells
Word to keep tracking edits the *reader* makes. It is independent of authoring
marks below (those work regardless).

```python
from docx_plus.revisions import enable_track_changes, disable_track_changes

enable_track_changes(doc)    # idempotent
disable_track_changes(doc)   # idempotent; does not touch existing marks
```

## Authoring insertions and deletions

Wrap run(s) **already in the document**. Target shapes match `add_comment`: a
single `Run`, a whole `Paragraph` (≥1 run), or a `(start_run, end_run)` tuple —
but a range must stay within one paragraph (`w:ins` / `w:del` cannot span a
paragraph boundary).

```python
from docx_plus.revisions import mark_insertion, mark_deletion, RevisionIdRegistry

p = doc.add_paragraph()
p.add_run("The plan ")
ins = p.add_run("ships in Q3 ")
old = p.add_run("ships someday")

reg = RevisionIdRegistry(doc)                       # share across a batch
mark_insertion(ins, author="Alice", id_registry=reg)
mark_deletion(old,  author="Bob",   id_registry=reg)   # retags w:t -> w:delText
```

`mark_insertion(target, *, author="", date=None, id_registry=None) -> RevisionRef`
and `mark_deletion(...)` (same signature). `RevisionRef` has `revision_id` and
`body_element`. `date` defaults to now (UTC); pass a `datetime` to override.

**All revision types share ONE id namespace** — a `w:ins` id and a `w:del` id
cannot collide. Use one `RevisionIdRegistry` per editing session.

## Reading revisions

```python
from docx_plus.revisions import read_revisions

for rv in read_revisions(doc):
    print(rv.revision_type, rv.author, repr(rv.text), rv.paragraph_index)
```

`read_revisions(doc) -> list[TrackedChange]`. `TrackedChange` is a frozen
dataclass: `revision_id`, `revision_type`, `author`, `timestamp`, `text`,
`paragraph_index`. `revision_type` is one of `insertion`, `deletion`,
`move_from`, `move_to`, `format_run`, `format_paragraph`,
`paragraph_mark_insertion`, `paragraph_mark_deletion`. Insertion text comes from
`<w:t>`, deletion text from `<w:delText>`; format and paragraph-mark changes have
empty `text`.

## Accepting / rejecting

```python
from docx_plus.revisions import (
    accept_revision, reject_revision,
    accept_all_revisions, reject_all_revisions,
)

accept_revision(doc, rev_id)   # keep the edit; RevisionNotFoundError (KeyError) if absent
reject_revision(doc, rev_id)   # restore pre-edit state
accept_all_revisions(doc)      # idempotent; resolves everything
reject_all_revisions(doc)
```

- **Insertion**: accept = keep the text (unwrap); reject = drop it.
- **Deletion**: accept = remove the text; reject = restore it as live `<w:t>`.
- **Move / format-change**: safe mechanical transforms.
- **Paragraph-mark revisions**: resolved with a non-corrupting fallback (the mark
  is dropped, text is left intact); true paragraph merge/split is deferred.

`*_all` processes innermost marks first, so nested revisions resolve cleanly.

## End-to-end

```python
from docx import Document
from docx_plus.revisions import (
    enable_track_changes, mark_insertion, mark_deletion,
    read_revisions, accept_all_revisions,
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
print(reopened.paragraphs[0].text)   # -> "Keep this new bit "
```

Type alias: `RevisionTarget = Run | Paragraph | tuple[Run, Run]`.

See also: `docx_plus/examples/track_changes.py`.
