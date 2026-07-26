"""Core foundation: namespaces, XML helpers, ID registry, package parts.

This subpackage is the only dependency target permitted to capability modules
(``styles/``, ``controls/``, ``fields/``, ``protection/``). See SPEC §9.1.

The submodules' public symbols are re-exported here so callers can use the
short form ``from docx_plus.core import IdRegistry, qn, el`` documented in
``docs/API.md`` — without losing access to the long form
``from docx_plus.core.ids import IdRegistry`` for code that wants to be
explicit about where a symbol lives.
"""

from docx_plus.core.errors import DocxPlusError
from docx_plus.core.ids import DuplicateIdError, IdRangeError, IdRegistry, ParaIdRegistry
from docx_plus.core.ns import MC, NSMAP, W14, W15, XML, A, InvalidNamespaceError, R, W, qn
from docx_plus.core.oxml import (
    body_document_for,
    build_complex_field,
    el,
    insert_before_first_anchor,
    remove,
    sub,
    xpath,
)
from docx_plus.core.parts import (
    COMMENTS_EXTENDED_SPEC,
    COMMENTS_SPEC,
    CT_COMMENTS_EXTENDED,
    ENDNOTES_SPEC,
    FOOTNOTES_SPEC,
    RT_COMMENTS_EXTENDED,
    PartSpec,
    get_or_create_part,
)

__all__ = [
    "A",
    "COMMENTS_EXTENDED_SPEC",
    "COMMENTS_SPEC",
    "CT_COMMENTS_EXTENDED",
    "ENDNOTES_SPEC",
    "FOOTNOTES_SPEC",
    "MC",
    "NSMAP",
    "R",
    "RT_COMMENTS_EXTENDED",
    "W",
    "W14",
    "W15",
    "XML",
    "DocxPlusError",
    "DuplicateIdError",
    "IdRangeError",
    "IdRegistry",
    "InvalidNamespaceError",
    "ParaIdRegistry",
    "PartSpec",
    "body_document_for",
    "build_complex_field",
    "el",
    "get_or_create_part",
    "insert_before_first_anchor",
    "qn",
    "remove",
    "sub",
    "xpath",
]
