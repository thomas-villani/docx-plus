# Footnotes and endnotes

`notes/write.py` exposes `add_footnote` and `add_endnote`, both with
identical shape: append a reference marker run to the paragraph, then
append a content entry in the corresponding [separate part](parts.md).
The content entry uses Word's `FootnoteText` / `EndnoteText` paragraph
style and `FootnoteReference` / `EndnoteReference` run style for the
leading reference glyph. The body text run carries
`xml:space="preserve"`.

For the calls, see the [notes guide](../guides/notes.md).

`edit_footnote(doc, id, text)` and `edit_endnote(doc, id, text)` mutate
the body of an existing note in place. They strip every `<w:p>` child
of the matching `<w:footnote>` / `<w:endnote>` element and append a
fresh paragraph built by the shared `_build_note_paragraph` helper
(used by both add and edit paths). The body-side reference marker in
the main document body is untouched, so the in-text superscript stays
put. Reserved separator ids (`-1`, `0`) raise `ValueError`; missing
ids raise `NoteNotFoundError`.

`read_footnotes(doc)` and `read_endnotes(doc)` walk the corresponding
part and pair each note with the paragraph index of its body-side
reference marker. Reserved entries (ids `-1` for separator, `0` for
continuation separator, or any entry with `w:type` of `"separator"` /
`"continuationSeparator"`) are filtered out before results are
returned, so callers only ever see user-authored notes.

`FootnoteIdRegistry` and `EndnoteIdRegistry` are two more disjoint
namespaces. The shared `_NoteIdRegistryBase` (`notes/registry.py`)
parameterises the relationship type and the note tag; the underlying
`_IdRegistryBase.reserve(value)` rejects values outside `[1, 2**31 - 1]`
on a range check, so ids `0` and `-1` are unissuable — the range check
fires before any duplicate check, so no special pre-seeding is
needed.
