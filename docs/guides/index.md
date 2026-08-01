# Guides

One page per capability, organised around what you are trying to do. Each
guide gives you working code first, then the traps, then links to the
[concept page](../concepts/index.md) explaining why the format behaves that
way and to the [module reference](../reference/core-ns.md) for full
signatures.

New here? Read [Getting started](../getting-started.md) first — it covers
the conventions every one of these guides assumes.

## Which guide do I need?

| I want to… | Guide |
|---|---|
| Find out why a paragraph looks the way it does; create, apply, or change styles; audit a whole document's formatting | [Styles and the cascade](styles.md) |
| Build a fillable form with text boxes, dropdowns, date pickers, checkboxes; read the values back; lock the document | [Forms and protection](forms.md) |
| Leave review comments anchored to specific text; reply to them; resolve and reopen threads | [Comments](comments.md) |
| Mark insertions and deletions as tracked changes; read revisions; accept or reject them | [Tracked changes](revisions.md) |
| Add a table of contents, numbered figure/table captions, or a table of figures | [Publishing](publishing.md) |
| Add footnotes or endnotes, and edit them later | [Footnotes and endnotes](notes.md) |
| Bookmark a heading and point a cross-reference at it | [Bookmarks and cross-references](bookmarks.md) |
| Insert page numbers, dates, or a raw field; make Word recalculate them | [Fields](fields.md) |
| Set up multi-column sections, mid-document section breaks, distinct even/odd headers, line numbers, page borders | [Page layout](layout.md) |
| Define bullet or multi-level numbered lists, apply them, restart numbering | [Lists and numbering](numbering.md) |
| Put borders and shading on tables; merge and unmerge cells | [Tables](tables.md) |
| Audit an inherited document for formatting defects, and plan the repair | [Linting](linting.md) |

## Common combinations

A few things need more than one module. These are the ones people hit:

- **A report with a working TOC** — [publishing](publishing.md) for the
  TOC and captions, [styles](styles.md) to make sure the headings carry
  real outline levels, and one `mark_fields_dirty` call from
  [fields](fields.md) before you save.
- **A fillable, locked form** — [forms](forms.md) covers both halves:
  `FormBuilder` to place the controls and `protect_document(mode="forms")`
  to stop readers editing anything else.
- **A review pass on someone else's document** — [linting](linting.md) to
  find what's wrong, [styles](styles.md) to reconcile the style ids, and
  [comments](comments.md) or [tracked changes](revisions.md) to record what
  you changed.
- **A legal or contract document** — [layout](layout.md) for line
  numbering, [notes](notes.md) for footnotes,
  [bookmarks](bookmarks.md) for clause cross-references.
