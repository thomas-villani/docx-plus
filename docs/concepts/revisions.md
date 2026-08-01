# Tracked changes

`revisions/` closes a gap python-docx cannot reach at all: it neither
reads nor writes tracked changes. The capability works entirely in
inline revision elements (`<w:ins>`, `<w:del>` with `<w:delText>`, the
`<w:moveFrom>` / `<w:moveTo>` move wrappers, and the property-change
markers) — none of which need a [separate part](parts.md), unlike
comments or notes.

For the calls, see the [tracked changes guide](../guides/revisions.md).

`revisions/mark.py:mark_insertion(target, ...)` wraps existing run(s)
in `<w:ins w:id w:author w:date>`, leaving each `<w:t>` untouched.
`mark_deletion(target, ...)` wraps the span in `<w:del>` and retags
every `<w:t>` in it to `<w:delText>` (the element Word uses for deleted
run text). Both take the same target shapes as comments — a `Run`, a
`Paragraph` (its first-to-last run span), or a `(start_run, end_run)`
tuple within one paragraph — and return a `RevisionRef` carrying the
assigned id and the wrapper element. `date=None` stamps the current
UTC time at millisecond precision.

`read_revisions(doc)` walks the body and returns one `TrackedChange`
per revision element in document order, dispatching on element tag to
classify each as one of the `RevisionType` literals (insertion,
deletion, move source/destination, run- or paragraph-property change,
paragraph-mark insertion/deletion). Each result pairs the affected
text with its `paragraph_index`. Move *range markers* are not reported
as separate entries — the wrapper that carries the moved text is.

`accept_revision(doc, id)` and `reject_revision(doc, id)` resolve a
single revision into final text — accepting an insertion unwraps it,
accepting a deletion removes it, and rejecting does the inverse — with
`accept_all_revisions` / `reject_all_revisions` as the bulk forms.
A missing id raises `RevisionNotFoundError`.

`enable_track_changes(doc)` / `disable_track_changes(doc)` toggle the
document-wide `<w:trackChanges>` flag in `settings.xml`, the switch
that makes Word record subsequent edits as revisions. It does **not**
retroactively mark anything the library already wrote — that is what
`mark_insertion` / `mark_deletion` are for.

`RevisionIdRegistry` owns the shared revision-id namespace, disjoint
from the SDT, comment, bookmark, and note namespaces ([invariant
3](invariants.md#the-invariants)). `mark_*` accept an `id_registry` to
share across an editing session, or build one scoped to the call from
the target's document.
