"""Document-level protection — the thing that enforces form-fill mode.

Adding content controls makes a document *fillable*, but it does not stop a
reader from editing surrounding paragraphs. ``w:documentProtection`` in
``settings.xml`` flips that switch: with ``mode="forms"`` Word locks every
range outside of an SDT, so the only thing the reader can edit is the
content-control fields.

v0.1 protection is **unpassworded**: it prevents accidental editing, not a
determined user. Password-protected forms (legacy hash algorithm; SPEC §1
non-goal) are deferred to v0.2.

This module imports only from ``docx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Literal

from docx_plus.core.ns import qn
from docx_plus.core.oxml import el, insert_before_first_anchor

if TYPE_CHECKING:
    from docx.document import Document as DocxDocument


ProtectionMode = Literal["forms", "readOnly", "comments", "trackedChanges"]


# Schema-strict insertion: in ``CT_Settings`` (ECMA-376 17.15.1.78),
# ``w:documentProtection`` sits between ``w:doNotTrackFormatting`` and
# ``w:autoFormatOverride``, well before ``w:defaultTabStop`` (SPEC §8 anchor).
# We insert before the first matching anchor we find.
_DOC_PROTECTION_LATER_SIBLINGS: tuple[str, ...] = (
    "w:autoFormatOverride",
    "w:styleLockTheme",
    "w:styleLockQFSet",
    "w:defaultTabStop",
    "w:autoHyphenation",
    "w:consecutiveHyphenLimit",
    "w:hyphenationZone",
    "w:doNotHyphenateCaps",
    "w:compat",
    "w:rsids",
    "w:themeFontLang",
)


def protect_document(
    doc: DocxDocument,
    *,
    mode: ProtectionMode = "forms",
) -> None:
    """Enforce document protection.

    Writes (or replaces) ``<w:documentProtection w:edit="MODE"
    w:enforcement="1"/>`` in ``settings.xml``. Idempotent: a second call with
    a different ``mode`` replaces the previous protection rather than
    stacking.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to protect.
        mode: What kind of editing to enforce.

            * ``"forms"`` (default) — only content controls are editable.
              Pair with :class:`docx_plus.controls.FormBuilder` to produce a
              fillable form.
            * ``"readOnly"`` — entire document is read-only.
            * ``"comments"`` — readers may only add comments.
            * ``"trackedChanges"`` — readers may edit with revisions on.

    Example:
        >>> from docx import Document
        >>> from docx_plus.protection import protect_document
        >>> doc = Document()
        >>> protect_document(doc, mode="forms")
    """
    settings = doc.settings.element
    existing = settings.find(qn("w:documentProtection"))
    if existing is not None:
        existing.set(qn("w:edit"), mode)
        existing.set(qn("w:enforcement"), "1")
        return
    new = el("w:documentProtection", **{"w:edit": mode, "w:enforcement": "1"})
    insert_before_first_anchor(settings, new, _DOC_PROTECTION_LATER_SIBLINGS)


def unprotect_document(doc: DocxDocument) -> None:
    """Remove any document protection.

    Idempotent: a no-op if the document is already unprotected.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to unlock.
    """
    settings = doc.settings.element
    existing = settings.find(qn("w:documentProtection"))
    if existing is not None:
        settings.remove(existing)


def is_protected(doc: DocxDocument) -> bool:
    """Return ``True`` if any protection is currently enforced.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to inspect.

    Returns:
        ``True`` if a ``w:documentProtection`` element is present in
        ``settings.xml``, ``False`` otherwise. Does not distinguish modes —
        read the element's ``w:edit`` attribute for that.
    """
    settings = doc.settings.element
    return settings.find(qn("w:documentProtection")) is not None


__all__ = [
    "ProtectionMode",
    "is_protected",
    "protect_document",
    "unprotect_document",
]
