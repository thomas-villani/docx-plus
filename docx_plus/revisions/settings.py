"""Toggle document-wide track-changes mode in ``settings.xml``.

Word records "track changes is on" with a single ``<w:trackChanges/>``
element in ``settings.xml``. Its presence is the flag — when it is
present (and not explicitly ``w:val="false"``) Word wraps every edit the
user makes in revision marks. :func:`mark_insertion` / :func:`mark_deletion`
write revision marks regardless of this flag; this toggle controls whether
*Word itself* keeps tracking once the user opens the document.

The element lives in a schema-ordered position in ``CT_Settings``
(ECMA-376 17.15.1.78): just after ``w:revisionView`` and before
``w:doNotTrackMoves`` / ``w:documentProtection`` / ``w:defaultTabStop``.
We insert before the first later sibling that exists (the
:func:`~docx_plus.core.oxml.insert_before_first_anchor` pattern proven in
``fields/update.py``); a wrong position makes Word and LibreOffice
silently drop the element.

This module imports only from ``docx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx_plus.core.ns import qn
from docx_plus.core.oxml import el, insert_before_first_anchor, remove

if TYPE_CHECKING:
    from docx.document import Document


# Schema-strict insertion: in ``CT_Settings`` ``w:trackChanges`` sits in the
# mid group right after ``w:revisionView``. We insert before the first of
# these later siblings we find; if none exist we fall back to appending.
# ``w:defaultTabStop`` is present in python-docx's default ``settings.xml``,
# so the common case lands here.
_TRACK_CHANGES_LATER_SIBLINGS: tuple[str, ...] = (
    "w:doNotTrackMoves",
    "w:doNotTrackFormatting",
    "w:documentProtection",
    "w:autoFormatOverride",
    "w:styleLockTheme",
    "w:styleLockQFSet",
    "w:defaultTabStop",
    "w:autoHyphenation",
    "w:consecutiveHyphenLimit",
    "w:hyphenationZone",
    "w:doNotHyphenateCaps",
    "w:characterSpacingControl",
    "w:defaultTableStyle",
    "w:evenAndOddHeaders",
    "w:compat",
    "w:rsids",
    "w:mathPr",
    "w:themeFontLang",
    "w:clrSchemeMapping",
    "w:shapeDefaults",
    "w:decimalSymbol",
    "w:listSeparator",
    "w:hdrShapeDefaults",
    "w:footnotePr",
    "w:endnotePr",
)


def enable_track_changes(doc: Document) -> None:
    """Turn on document-wide track-changes mode.

    Writes ``<w:trackChanges/>`` into ``settings.xml`` so Word records every
    subsequent user edit as a revision. Idempotent: calling twice leaves a
    single element. A pre-existing element is normalised to "on" by
    stripping any ``w:val="false"`` attribute; should a malformed
    ``settings.xml`` carry several copies, they collapse to one
    (``CT_Settings`` permits at most one).

    Args:
        doc: The python-docx :class:`~docx.document.Document` whose settings
            part should be flagged.

    Example:
        >>> from docx import Document
        >>> from docx_plus.revisions import enable_track_changes
        >>> doc = Document()
        >>> enable_track_changes(doc)
    """
    settings = doc.settings.element
    existing = settings.findall(qn("w:trackChanges"))
    if existing:
        existing[0].attrib.pop(qn("w:val"), None)
        for extra in existing[1:]:
            remove(extra)
        return
    insert_before_first_anchor(settings, el("w:trackChanges"), _TRACK_CHANGES_LATER_SIBLINGS)


def disable_track_changes(doc: Document) -> None:
    """Turn off document-wide track-changes mode.

    Removes every ``<w:trackChanges/>`` from ``settings.xml``. Idempotent:
    a document that was never tracking is left unchanged. Existing revision
    marks already in the body are untouched — use
    :func:`~docx_plus.revisions.accept_all_revisions` /
    :func:`~docx_plus.revisions.reject_all_revisions` to resolve those.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to update.
    """
    settings = doc.settings.element
    for elem in settings.findall(qn("w:trackChanges")):
        remove(elem)


__all__ = ["disable_track_changes", "enable_track_changes"]
