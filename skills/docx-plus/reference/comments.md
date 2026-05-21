# Comments — anchored review comments

Module: `docx_plus.comments`. Add, read, edit, and delete Word review comments
that are **anchored** to specific text.

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

`add_comment(target, text, *, author="", initials=None, id_registry=None) -> CommentRef`
(`CommentRef` has `comment_id`, `body_element`).

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

## Reading comments

```python
from docx_plus.comments import read_comments

for c in read_comments(doc):
    print(f"[{c.author}] {c.text!r} on {c.anchored_text!r} (p{c.paragraph_index})")
```

`read_comments(doc) -> list[AnchoredComment]`. An `AnchoredComment` is a frozen
dataclass with `comment_id`, `author`, `initials`, `timestamp`, `text`,
`anchored_text` (the document text the comment is anchored to), and
`paragraph_index`.

## Editing and deleting

```python
from docx_plus.comments import edit_comment, delete_comment, clear_all_comments

edit_comment(doc, ref.comment_id, "Revised note.")  # body text only; keeps
                                                     # author/date/initials/anchors
delete_comment(doc, ref.comment_id)   # removes all four traces; idempotent
clear_all_comments(doc)               # delete every comment; idempotent
```

- `edit_comment(doc, comment_id, text)` — replaces body text in place; preserves
  `w:author` / `w:date` / `w:initials` and the body anchors. Raises
  `CommentNotFoundError` (`KeyError`) on an unknown id.
- `delete_comment(doc, comment_id)` — removes the range markers, the reference
  run, and the body. Idempotent (unknown id is a no-op).
- `clear_all_comments(doc)` — routes every id through `delete_comment`.

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

See also: `docx_plus/examples/add_comments.py`.
