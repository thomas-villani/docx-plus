"""Tracked changes — read, author, and resolve OOXML revision marks.

python-docx cannot read or write tracked changes at all. This package
fills that gap with the inline revision elements Word uses — ``w:ins``,
``w:del`` (with ``w:delText``), the move wrappers, and the property-change
markers — none of which need a separate part (unlike comments).

Public surface:

- :func:`enable_track_changes` / :func:`disable_track_changes` — toggle the
  document-wide ``w:trackChanges`` flag in ``settings.xml``
- :func:`mark_insertion` / :func:`mark_deletion` — wrap existing run(s) as a
  tracked insertion or deletion
- :func:`read_revisions` — enumerate every revision with its metadata and
  affected text
- :func:`accept_revision` / :func:`reject_revision` and the
  :func:`accept_all_revisions` / :func:`reject_all_revisions` bulk forms —
  resolve revisions into final text
- :class:`RevisionRef` — write-side handle; :class:`TrackedChange` — read-side
  result; :class:`RevisionType` — the revision-type literal
- :class:`RevisionIdRegistry` — share across an editing session for many marks
- :class:`RevisionNotFoundError` — raised by accept/reject for a missing id

See `ROADMAP.md` §1 for where this capability was scoped.
"""

from __future__ import annotations

from docx_plus.revisions.accept import (
    accept_all_revisions,
    accept_revision,
    reject_all_revisions,
    reject_revision,
)
from docx_plus.revisions.mark import (
    RevisionNotFoundError,
    RevisionRef,
    RevisionTarget,
    mark_deletion,
    mark_insertion,
)
from docx_plus.revisions.read import RevisionType, TrackedChange, read_revisions
from docx_plus.revisions.registry import RevisionIdRegistry
from docx_plus.revisions.settings import disable_track_changes, enable_track_changes

__all__ = [
    "RevisionIdRegistry",
    "RevisionNotFoundError",
    "RevisionRef",
    "RevisionTarget",
    "RevisionType",
    "TrackedChange",
    "accept_all_revisions",
    "accept_revision",
    "disable_track_changes",
    "enable_track_changes",
    "mark_deletion",
    "mark_insertion",
    "read_revisions",
    "reject_all_revisions",
    "reject_revision",
]
