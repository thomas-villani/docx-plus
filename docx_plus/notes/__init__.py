"""Footnotes and endnotes — insert-only API for v0.2.

Both note kinds live in separate OOXML parts (``word/footnotes.xml``,
``word/endnotes.xml``) created on first use. The body of the main
document holds the reference marker (a small superscripted glyph in
Word's rendering); the note text lives in the corresponding part.

Public surface:

- :func:`add_footnote` / :func:`add_endnote` — insert at a paragraph
- :func:`read_footnotes` / :func:`read_endnotes` — list user-authored
  notes (separator entries are filtered)
- :class:`NoteContent` — read-side result
- :class:`FootnoteRef` / :class:`EndnoteRef` — write-side handles
- :class:`FootnoteIdRegistry` / :class:`EndnoteIdRegistry` — share
  across an editing session

See SPEC §15 (the post-v0.1 roadmap) for where this capability was scoped.
"""

from __future__ import annotations

from docx_plus.notes.read import NoteContent, read_endnotes, read_footnotes
from docx_plus.notes.registry import EndnoteIdRegistry, FootnoteIdRegistry
from docx_plus.notes.write import (
    EndnoteRef,
    FootnoteRef,
    NoteNotFoundError,
    add_endnote,
    add_footnote,
    edit_endnote,
    edit_footnote,
)

__all__ = [
    "EndnoteIdRegistry",
    "EndnoteRef",
    "FootnoteIdRegistry",
    "FootnoteRef",
    "NoteContent",
    "NoteNotFoundError",
    "add_endnote",
    "add_footnote",
    "edit_endnote",
    "edit_footnote",
    "read_endnotes",
    "read_footnotes",
]
