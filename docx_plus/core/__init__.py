"""Core foundation: namespaces, XML helpers, ID registry, package parts.

This subpackage is the only dependency target permitted to capability modules
(``styles/``, ``controls/``, ``fields/``, ``protection/``). See SPEC §9.1.

The submodules' public symbols are re-exported here so callers can use the
short form ``from docx_plus.core import IdRegistry, qn, el`` documented in
``docs/API.md`` — without losing access to the long form
``from docx_plus.core.ids import IdRegistry`` for code that wants to be
explicit about where a symbol lives.
"""

from docx_plus.core.borders import Border, border_attrs
from docx_plus.core.errors import DocxPlusError
from docx_plus.core.ids import (
    BookmarkIdRegistry,
    BookmarkNameRegistry,
    DuplicateBookmarkNameError,
    DuplicateIdError,
    IdRangeError,
    IdRegistry,
    ParaIdRegistry,
)
from docx_plus.core.ns import (
    MC,
    NSMAP,
    W14,
    W15,
    W16CID,
    XML,
    A,
    InvalidNamespaceError,
    R,
    W,
    qn,
)
from docx_plus.core.oxml import (
    body_document_for,
    build_bookmark,
    build_complex_field,
    el,
    insert_before_first_anchor,
    ordered_insert,
    remove,
    sub,
    validate_bookmark_name,
    xpath,
)
from docx_plus.core.parts import (
    COMMENTS_EXTENDED_SPEC,
    COMMENTS_IDS_SPEC,
    COMMENTS_SPEC,
    CT_COMMENTS_EXTENDED,
    CT_COMMENTS_IDS,
    CT_PEOPLE,
    ENDNOTES_SPEC,
    FOOTNOTES_SPEC,
    NUMBERING_SPEC,
    PEOPLE_SPEC,
    RT_COMMENTS_EXTENDED,
    RT_COMMENTS_IDS,
    RT_PEOPLE,
    PartSpec,
    get_or_create_part,
)

__all__ = [
    "A",
    "COMMENTS_EXTENDED_SPEC",
    "COMMENTS_IDS_SPEC",
    "COMMENTS_SPEC",
    "CT_COMMENTS_EXTENDED",
    "CT_COMMENTS_IDS",
    "CT_PEOPLE",
    "ENDNOTES_SPEC",
    "FOOTNOTES_SPEC",
    "MC",
    "NSMAP",
    "NUMBERING_SPEC",
    "PEOPLE_SPEC",
    "R",
    "RT_COMMENTS_EXTENDED",
    "RT_COMMENTS_IDS",
    "RT_PEOPLE",
    "W",
    "W14",
    "W15",
    "W16CID",
    "XML",
    "BookmarkIdRegistry",
    "BookmarkNameRegistry",
    "Border",
    "DocxPlusError",
    "DuplicateBookmarkNameError",
    "DuplicateIdError",
    "IdRangeError",
    "IdRegistry",
    "InvalidNamespaceError",
    "ParaIdRegistry",
    "PartSpec",
    "body_document_for",
    "border_attrs",
    "build_bookmark",
    "build_complex_field",
    "el",
    "get_or_create_part",
    "insert_before_first_anchor",
    "ordered_insert",
    "qn",
    "remove",
    "sub",
    "validate_bookmark_name",
    "xpath",
]
