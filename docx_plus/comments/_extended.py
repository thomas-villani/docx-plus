"""``commentsExtended.xml`` plumbing — the thread graph behind comments.

Word 2013 introduced *threaded* comments, but it did not extend
``<w:comment>`` to carry the thread. Instead it added a second part,
``/word/commentsExtended.xml``, holding one ``<w15:commentEx>`` per
comment:

.. code-block:: xml

    <w15:commentsEx>
      <w15:commentEx w15:paraId="3F2A19C4" w15:done="0"/>
      <w15:commentEx w15:paraId="5B71E0A2"
                     w15:paraIdParent="3F2A19C4" w15:done="0"/>
    </w15:commentsEx>

Two properties of that design drive everything in this module:

1. **Entries key off ``w14:paraId``, not ``w:id``.** The key is the
   ``paraId`` of the *last paragraph* of the comment body. Comment ids
   and thread keys are therefore separate namespaces, and every mapping
   between them has to go through the comment body.
2. **The part is optional.** A document whose comments predate Word 2013
   — or that came from python-docx — has no extended part at all. Every
   reader here treats that as "one unresolved root per comment" rather
   than an error.

Internal module: nothing here is re-exported from
:mod:`docx_plus.comments`. :mod:`docx_plus.comments.anchor` and
:mod:`docx_plus.comments.threads` share it so the two write paths cannot
drift.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import XmlPart
from lxml import etree

from docx_plus.core.ids import ParaIdRegistry
from docx_plus.core.ns import qn
from docx_plus.core.oxml import el, remove, xpath
from docx_plus.core.parts import (
    COMMENTS_EXTENDED_SPEC,
    RT_COMMENTS_EXTENDED,
    get_or_create_part,
)

if TYPE_CHECKING:
    from docx.document import Document


# ---------------------------------------------------------------------------
# paraId stamping — the bridge between a comment body and its thread entry.
# ---------------------------------------------------------------------------


def stamp_para_ids(comment_el: etree._Element, registry: ParaIdRegistry) -> str:
    """Ensure every paragraph in a comment body carries a ``w14:paraId``.

    Paragraphs that already have one keep it — re-stamping would break
    the ``w15:commentEx`` entry pointing at the old value, silently
    flattening the thread.

    Args:
        comment_el: A ``<w:comment>`` element.
        registry: Allocator for fresh ``paraId`` values.

    Returns:
        The thread key — the ``paraId`` of the comment's *last*
        paragraph, which is what ``commentsExtended.xml`` references.
        Empty string for a comment with no paragraphs at all (malformed,
        but not worth raising over: the caller simply writes no entry).
    """
    attr = qn("w14:paraId")
    paragraphs = xpath(comment_el, ".//w:p")
    if not paragraphs:
        return ""

    for paragraph in paragraphs:
        if not paragraph.get(attr):
            paragraph.set(attr, registry.next_hex())

    return paragraphs[-1].get(attr) or ""


def thread_key(comment_el: etree._Element) -> str | None:
    """Return a comment body's thread key without allocating anything.

    Args:
        comment_el: A ``<w:comment>`` element.

    Returns:
        The ``w14:paraId`` of the last paragraph, or ``None`` if the
        comment has no paragraphs or the last one is unstamped.
    """
    paragraphs = xpath(comment_el, ".//w:p")
    if not paragraphs:
        return None
    return cast("str | None", paragraphs[-1].get(qn("w14:paraId")))


# ---------------------------------------------------------------------------
# The extended part itself.
# ---------------------------------------------------------------------------


def extended_root(doc: Document) -> etree._Element | None:
    """Return the ``<w15:commentsEx>`` root, or ``None`` if absent.

    Read-only counterpart to :func:`get_or_create_extended_root` — it
    never fabricates the part, so callers that only inspect threads do
    not add one to documents that had none.
    """
    try:
        part = cast("XmlPart", doc.part.part_related_by(RT_COMMENTS_EXTENDED))
    except KeyError:
        return None
    return cast("etree._Element", part.element)


def get_or_create_extended_root(doc: Document) -> etree._Element:
    """Return the ``<w15:commentsEx>`` root, creating the part on first use."""
    _, root = get_or_create_part(doc, COMMENTS_EXTENDED_SPEC)
    return root


def find_comment_ex(root: etree._Element, para_id: str) -> etree._Element | None:
    """Return the ``<w15:commentEx>`` for ``para_id``, or ``None``."""
    matches = xpath(root, "./w15:commentEx[@w15:paraId=$pid]", pid=para_id)
    return matches[0] if matches else None


def upsert_comment_ex(
    doc: Document,
    para_id: str,
    *,
    parent_para_id: str | None = None,
    done: bool | None = None,
) -> etree._Element | None:
    """Create or update the thread entry for ``para_id``.

    Args:
        doc: Document whose extended part is written.
        para_id: Thread key — the comment's last-paragraph ``paraId``.
            An empty value is a no-op (see :func:`stamp_para_ids`).
        parent_para_id: Thread key of the parent comment for a reply, or
            ``None`` to leave the existing parentage untouched. Pass
            ``""`` to explicitly clear it and promote the comment to a
            thread root.
        done: Resolved flag. ``None`` leaves an existing entry's flag
            alone; a newly created entry defaults to unresolved.

    Returns:
        The ``<w15:commentEx>`` element, or ``None`` when ``para_id`` is
        empty.
    """
    if not para_id:
        return None

    root = get_or_create_extended_root(doc)
    entry = find_comment_ex(root, para_id)
    if entry is None:
        entry = el("w15:commentEx", **{"w15:paraId": para_id})
        root.append(entry)
        if done is None:
            done = False

    if parent_para_id is not None:
        parent_attr = qn("w15:paraIdParent")
        if parent_para_id:
            entry.set(parent_attr, parent_para_id)
        elif parent_attr in entry.attrib:
            del entry.attrib[parent_attr]

    if done is not None:
        entry.set(qn("w15:done"), "1" if done else "0")

    return entry


def drop_comment_ex(doc: Document, para_id: str) -> None:
    """Remove the thread entry for ``para_id``. Idempotent."""
    root = extended_root(doc)
    if root is None or not para_id:
        return
    for entry in xpath(root, "./w15:commentEx[@w15:paraId=$pid]", pid=para_id):
        remove(entry)


def clear_comments_ex(doc: Document, *, remove_part: bool = False) -> None:
    """Remove every thread entry, optionally tearing down the part.

    Mirrors the ``remove_part`` contract of
    :func:`docx_plus.comments.clear_all_comments`: by default the empty
    part is kept so a subsequent insert reuses the relationship.
    """
    root = extended_root(doc)
    if root is None:
        return

    if remove_part:
        for rid, rel in list(doc.part.rels.items()):
            if rel.reltype == RT_COMMENTS_EXTENDED:
                doc.part.drop_rel(rid)
        return

    for entry in xpath(root, "./w15:commentEx"):
        remove(entry)


# ---------------------------------------------------------------------------
# Comment-id <-> thread-key mappings.
# ---------------------------------------------------------------------------


def comment_elements(doc: Document) -> dict[int, etree._Element]:
    """Return ``{comment_id: <w:comment>}`` for every comment in ``doc``.

    Comments whose ``w:id`` is missing or non-integer are skipped, in
    line with :func:`docx_plus.comments.read_comments`.
    """
    try:
        part = cast("XmlPart", doc.part.part_related_by(RT.COMMENTS))
    except KeyError:
        return {}

    result: dict[int, etree._Element] = {}
    for comment_el in xpath(part.element, "./w:comment"):
        raw = comment_el.get(qn("w:id"))
        if raw is None:
            continue
        try:
            result[int(raw)] = comment_el
        except ValueError:
            continue
    return result


def key_maps(doc: Document) -> tuple[dict[int, str], dict[str, int]]:
    """Return ``(comment_id -> thread key, thread key -> comment_id)``.

    Only comments whose body carries a ``w14:paraId`` appear; an
    unstamped comment cannot participate in a thread.
    """
    by_id: dict[int, str] = {}
    by_key: dict[str, int] = {}
    for comment_id, comment_el in comment_elements(doc).items():
        key = thread_key(comment_el)
        if not key:
            continue
        by_id[comment_id] = key
        by_key.setdefault(key, comment_id)
    return by_id, by_key


def thread_state(doc: Document) -> dict[int, tuple[int | None, bool]]:
    """Return ``{comment_id: (parent_comment_id, resolved)}`` for ``doc``.

    A comment with no ``<w15:commentEx>`` entry — the whole document, if
    the extended part is absent — reports as an unresolved root
    (``(None, False)``). A ``paraIdParent`` pointing at a key no comment
    owns is treated the same way: dangling parentage promotes the
    comment to a root rather than dropping it from the graph.
    """
    state: dict[int, tuple[int | None, bool]] = {}
    by_id, by_key = key_maps(doc)
    root = extended_root(doc)

    for comment_id in comment_elements(doc):
        state[comment_id] = (None, False)
    if root is None:
        return state

    entries: dict[str, etree._Element] = {}
    for entry in xpath(root, "./w15:commentEx"):
        key = entry.get(qn("w15:paraId"))
        if key:
            entries.setdefault(key, entry)

    for comment_id, key in by_id.items():
        entry = entries.get(key)
        if entry is None:
            continue
        parent_key = entry.get(qn("w15:paraIdParent"))
        parent_id = by_key.get(parent_key) if parent_key else None
        if parent_id == comment_id:
            parent_id = None  # self-parented entry: treat as a root
        state[comment_id] = (parent_id, _is_done(entry))

    return state


def root_id_of(state: dict[int, tuple[int | None, bool]], comment_id: int) -> int:
    """Walk ``comment_id`` up to its thread root.

    Word only ever nests one level deep — every reply points straight at
    the thread root — but ``commentsExtended.xml`` permits a chain, so
    this follows the parentage as far as it goes. A cycle in malformed
    parentage terminates at the first repeat rather than spinning.
    """
    seen = {comment_id}
    current = comment_id
    while True:
        parent = state.get(current, (None, False))[0]
        if parent is None or parent in seen:
            return current
        seen.add(parent)
        current = parent


def descendant_ids(state: dict[int, tuple[int | None, bool]], root_id: int) -> list[int]:
    """Return every reply beneath ``root_id``, breadth-first.

    Excludes ``root_id`` itself. Cycle-safe for the same reason as
    :func:`root_id_of`.
    """
    children: dict[int, list[int]] = {}
    for comment_id, (parent, _) in state.items():
        if parent is not None:
            children.setdefault(parent, []).append(comment_id)

    found: list[int] = []
    seen = {root_id}
    queue = list(children.get(root_id, ()))
    while queue:
        comment_id = queue.pop(0)
        if comment_id in seen:
            continue
        seen.add(comment_id)
        found.append(comment_id)
        queue.extend(children.get(comment_id, ()))
    return found


def _is_done(entry: etree._Element) -> bool:
    """Interpret ``w15:done`` per the OOXML boolean-attribute rules.

    ECMA-376 Part 1 §22.9.2.7 (``ST_OnOff``) accepts ``1``/``0``,
    ``true``/``false``, and ``on``/``off``. Word writes ``0`` and ``1``;
    other producers do not always.
    """
    raw = entry.get(qn("w15:done"))
    if raw is None:
        return False
    return raw.strip().lower() in {"1", "true", "on"}


__all__ = [
    "clear_comments_ex",
    "comment_elements",
    "descendant_ids",
    "drop_comment_ex",
    "extended_root",
    "find_comment_ex",
    "get_or_create_extended_root",
    "key_maps",
    "root_id_of",
    "stamp_para_ids",
    "thread_key",
    "thread_state",
    "upsert_comment_ex",
]
