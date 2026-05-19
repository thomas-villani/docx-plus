"""Core foundation: namespaces, XML helpers, ID registry, package parts.

This subpackage is the only dependency target permitted to capability modules
(``styles/``, ``controls/``, ``fields/``, ``protection/``). See SPEC §9.1.

The submodules' public symbols are re-exported here so callers can use the
short form ``from docx_plus.core import IdRegistry, qn, el`` documented in
``docs/API.md`` — without losing access to the long form
``from docx_plus.core.ids import IdRegistry`` for code that wants to be
explicit about where a symbol lives.
"""


class DocxPlusError(Exception):
    """Base class for all library-raised errors.

    Every typed error in docx_plus subclasses this so callers can catch the
    library's failures without catching unrelated ``ValueError``/``RuntimeError``
    instances from python-docx or lxml.
    """


from docx_plus.core.ids import DuplicateIdError, IdRangeError, IdRegistry  # noqa: E402
from docx_plus.core.ns import MC, NSMAP, W14, XML, A, InvalidNamespaceError, R, W, qn  # noqa: E402
from docx_plus.core.oxml import (  # noqa: E402
    build_complex_field,
    el,
    insert_before_first_anchor,
    remove,
    sub,
    xpath,
)
from docx_plus.core.parts import (  # noqa: E402
    COMMENTS_SPEC,
    ENDNOTES_SPEC,
    FOOTNOTES_SPEC,
    PartSpec,
    get_or_create_part,
)

__all__ = [
    "A",
    "COMMENTS_SPEC",
    "ENDNOTES_SPEC",
    "FOOTNOTES_SPEC",
    "MC",
    "NSMAP",
    "R",
    "W",
    "W14",
    "XML",
    "DocxPlusError",
    "DuplicateIdError",
    "IdRangeError",
    "IdRegistry",
    "InvalidNamespaceError",
    "PartSpec",
    "build_complex_field",
    "el",
    "get_or_create_part",
    "insert_before_first_anchor",
    "qn",
    "remove",
    "sub",
    "xpath",
]
