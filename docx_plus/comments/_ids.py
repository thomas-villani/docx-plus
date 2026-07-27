"""``commentsIds.xml`` plumbing — durable comment identity.

A comment has three identifiers and only one of them is stable:

- ``w:id`` — a position-dependent index in ``comments.xml``. Word
  renumbers it freely.
- ``w14:paraId`` — the body-paragraph id the thread graph keys off. It
  changes whenever the comment's *body* is rewritten.
- ``w16cid:durableId`` — added by Word 2016 for exactly this reason.
  Stable for the life of the comment.

Anything that needs to cite a comment across edits — a permalink, an
external review tracker, a diff between two revisions of a document —
needs the third.

.. code-block:: xml

    <w16cid:commentsIds>
      <w16cid:commentId w16cid:paraId="18B75E5F" w16cid:durableId="33EF1546"/>
    </w16cid:commentsIds>

Two things follow from that shape:

1. **Entries key off ``w14:paraId``**, exactly as ``commentsExtended.xml``
   does — so this module reuses :mod:`docx_plus.comments._extended`'s
   ``stamp_para_ids`` / ``thread_key`` / ``key_maps`` rather than
   building a second bridge from comment ids to part entries.
2. **``durableId`` is hex**, not decimal: ``ST_LongHexNumber``, the same
   8-uppercase-digit rendering as ``paraId``. Confirmed against a
   Word-authored file, which wrote ``33EF1546`` / ``31436C50`` /
   ``50E18CF9``.

The part is optional, and Word regenerates missing entries on open, so
every reader here treats an absent part as "no durable identity yet"
rather than an error.

Internal module: nothing here is re-exported from
:mod:`docx_plus.comments`. The durable id surfaces publicly as
``AnchoredComment.durable_id``.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.opc.part import XmlPart

from docx_plus.comments.registry import DurableIdRegistry
from docx_plus.core.ns import qn
from docx_plus.core.oxml import el, remove, xpath
from docx_plus.core.parts import (
    COMMENTS_IDS_SPEC,
    RT_COMMENTS_IDS,
    get_or_create_part,
)

if TYPE_CHECKING:
    from docx.document import Document
    from lxml import etree


def ids_root(doc: Document) -> etree._Element | None:
    """Return the ``<w16cid:commentsIds>`` root, or ``None`` if absent.

    Read-only counterpart to :func:`get_or_create_ids_root` — it never
    fabricates the part, so callers that only inspect durable ids do not
    add one to documents that had none.
    """
    try:
        part = cast("XmlPart", doc.part.part_related_by(RT_COMMENTS_IDS))
    except KeyError:
        return None
    return cast("etree._Element", part.element)


def get_or_create_ids_root(doc: Document) -> etree._Element:
    """Return the ``<w16cid:commentsIds>`` root, creating the part on first use."""
    _, root = get_or_create_part(doc, COMMENTS_IDS_SPEC)
    return root


def find_comment_id(root: etree._Element, para_id: str) -> etree._Element | None:
    """Return the ``<w16cid:commentId>`` for ``para_id``, or ``None``."""
    matches = xpath(root, "./w16cid:commentId[@w16cid:paraId=$pid]", pid=para_id)
    return matches[0] if matches else None


def upsert_comment_id(
    doc: Document,
    para_id: str,
    *,
    registry: DurableIdRegistry | None = None,
) -> str | None:
    """Ensure ``para_id`` has a durable id, minting one if needed.

    An existing entry keeps its ``durableId`` — reissuing would defeat
    the point of the part, breaking every reference already taken
    against the old value.

    Args:
        doc: Document whose ids part is written.
        para_id: Thread key — the comment's last-paragraph ``paraId``.
            An empty value is a no-op, matching
            :func:`docx_plus.comments._extended.upsert_comment_ex`.
        registry: Allocator to share across a batch of inserts. Built
            from ``doc`` when omitted.

    Returns:
        The durable id as 8 uppercase hex digits, or ``None`` when
        ``para_id`` is empty.
    """
    if not para_id:
        return None

    root = get_or_create_ids_root(doc)
    existing = find_comment_id(root, para_id)
    if existing is not None:
        return existing.get(qn("w16cid:durableId"))

    if registry is None:
        registry = DurableIdRegistry(doc)
    durable_id = registry.next_hex()
    root.append(
        el(
            "w16cid:commentId",
            **{"w16cid:paraId": para_id, "w16cid:durableId": durable_id},
        )
    )
    return durable_id


def drop_comment_id(doc: Document, para_id: str) -> None:
    """Remove the durable-id entry for ``para_id``. Idempotent."""
    root = ids_root(doc)
    if root is None or not para_id:
        return
    for entry in xpath(root, "./w16cid:commentId[@w16cid:paraId=$pid]", pid=para_id):
        remove(entry)


def clear_comment_ids(doc: Document, *, remove_part: bool = False) -> None:
    """Remove every durable-id entry, optionally tearing down the part.

    Mirrors the ``remove_part`` contract of
    :func:`docx_plus.comments.clear_all_comments`.
    """
    root = ids_root(doc)
    if root is None:
        return

    if remove_part:
        for rid, rel in list(doc.part.rels.items()):
            if rel.reltype == RT_COMMENTS_IDS:
                doc.part.drop_rel(rid)
        return

    for entry in xpath(root, "./w16cid:commentId"):
        remove(entry)


def durable_id_map(doc: Document) -> dict[str, str]:
    """Return ``{thread key: durable id}`` for every entry in ``doc``.

    Empty when the part is absent, which is the state of every document
    written before Word 2016 or by any other producer.
    """
    root = ids_root(doc)
    if root is None:
        return {}

    mapping: dict[str, str] = {}
    for entry in xpath(root, "./w16cid:commentId"):
        para_id = entry.get(qn("w16cid:paraId"))
        durable_id = entry.get(qn("w16cid:durableId"))
        if para_id and durable_id:
            mapping.setdefault(para_id, durable_id)
    return mapping


__all__ = [
    "clear_comment_ids",
    "drop_comment_id",
    "durable_id_map",
    "find_comment_id",
    "get_or_create_ids_root",
    "ids_root",
    "upsert_comment_id",
]
