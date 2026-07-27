"""Comment id registries — ``w:id`` and ``w16cid:durableId``.

Comment ``w:id`` lives in a separate uniqueness namespace from SDT ids,
bookmark ids, and note ids (a comment with id ``5`` does not collide
with a bookmark with id ``5``). This module ships a tiny subclass of
:class:`~docx_plus.core.ids._IdRegistryBase` that seeds itself from the
comment ids already present in ``comments.xml`` and from any orphaned
body-side markers (``w:commentRangeStart``, ``w:commentRangeEnd``,
``w:commentReference``) still in the body. The latter matters because
hand-edited or partially-deleted documents can leave any one of those
markers behind on its own, and it should still block id reuse.

This module imports only from ``docx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import XmlPart

from docx_plus.core.ids import _IdRegistryBase
from docx_plus.core.parts import RT_COMMENTS_IDS

if TYPE_CHECKING:
    from docx.document import Document


class CommentIdRegistry(_IdRegistryBase):
    """Tracks issued comment ``w:id`` values for one document-edit session."""

    def _seed_from_document(self, doc: Document) -> None:
        document_part = doc.part
        try:
            comments_part = cast("XmlPart", document_part.part_related_by(RT.COMMENTS))
        except KeyError:
            comments_root = None
        else:
            comments_root = comments_part.element

        if comments_root is not None:
            self._collect_id_attrs(comments_root, "./w:comment")

        # Body-side anchors — protect against orphaned ranges left by other
        # tools that wrote the range markers but skipped the comment body.
        # All three body-side elements carry the id, and any one of them can
        # survive on its own after a partial hand-edit (e.g. rangeStart
        # stripped but rangeEnd left behind), so all three must block reuse.
        body = doc.element.body
        self._collect_id_attrs(body, ".//w:commentRangeStart")
        self._collect_id_attrs(body, ".//w:commentRangeEnd")
        self._collect_id_attrs(body, ".//w:commentReference")


class DurableIdRegistry(_IdRegistryBase):
    """Tracks issued ``w16cid:durableId`` values for one edit session.

    A durable id is a comment's *stable* identity. ``w:id`` is a
    position-dependent index Word renumbers freely, and ``w14:paraId``
    changes whenever a comment's body is rewritten — neither survives as
    a citable handle. ``commentsIds.xml`` exists to give each comment one
    that does, which is what a permalink into a review thread needs.

    Word writes ``durableId`` as 8 uppercase hex digits, the same
    ``ST_LongHexNumber`` rendering as ``w14:paraId`` — verified against a
    Word-authored file, which produced values like ``33EF1546``. Use
    :meth:`~docx_plus.core.ids._IdRegistryBase.next_hex`, not
    :meth:`~docx_plus.core.ids._IdRegistryBase.next`.

    Seeds from ``commentsIds.xml`` only. Unlike ``paraId``, a durable id
    is scoped to that one part, so it cannot collide with an id in the
    body or in any other part.
    """

    def _seed_from_document(self, doc: Document) -> None:
        try:
            part = cast("XmlPart", doc.part.part_related_by(RT_COMMENTS_IDS))
        except KeyError:
            return
        self._collect_hex_id_attrs(part.element, "./w16cid:commentId", "w16cid:durableId")


__all__ = ["CommentIdRegistry", "DurableIdRegistry"]
