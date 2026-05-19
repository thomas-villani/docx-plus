"""Doc-level ``<w:evenAndOddHeaders>`` switch in ``settings.xml``.

This flag is constantly confused with
``<w:titlePg>`` (per-section, controls whether first-page differs) and
``Section.different_first_page_header_footer`` (the python-docx wrapper
around that). The doc-level flag here is different: it tells Word that
even-numbered pages may have a different header/footer from odd-numbered
pages, *across every section*. python-docx exposes no setter for it, so
this module writes the element directly into ``settings.xml`` using the
same schema-strict insertion pattern as
:func:`docx_plus.fields.mark_fields_dirty`.

This module imports only from ``docx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx_plus.core.ns import qn
from docx_plus.core.oxml import el, insert_before_first_anchor, remove

if TYPE_CHECKING:
    from docx.document import Document as DocxDocument


# CT_Settings schema order (ECMA-376 17.15.1.78). ``w:evenAndOddHeaders``
# is the third optional child; everything below comes after it. Listed
# in schema order so the search picks the first existing later-sibling.
_EVEN_AND_ODD_HEADERS_LATER_SIBLINGS: tuple[str, ...] = (
    "w:noPunctuationKerning",
    "w:characterSpacingControl",
    "w:printTwoOnOne",
    "w:strictFirstAndLastChars",
    "w:noLineBreaksAfter",
    "w:noLineBreaksBefore",
    "w:savePreviewPicture",
    "w:doNotValidateAgainstSchema",
    "w:saveInvalidXml",
    "w:ignoreMixedContent",
    "w:alwaysShowPlaceholderText",
    "w:doNotDemarcateInvalidXml",
    "w:saveXmlDataOnly",
    "w:useXSLTWhenSaving",
    "w:saveThroughXslt",
    "w:showXMLTags",
    "w:alwaysMergeEmptyNamespace",
    "w:updateFields",
    "w:hdrShapeDefaults",
    "w:footnotePr",
    "w:endnotePr",
    "w:compat",
    "w:docVars",
    "w:rsids",
    "w:mathPr",
    "w:themeFontLang",
    "w:clrSchemeMapping",
    "w:shapeDefaults",
    "w:decimalSymbol",
    "w:listSeparator",
)


def enable_distinct_even_odd_headers(doc: DocxDocument) -> None:
    """Enable distinct even-page vs odd-page headers and footers.

    Writes ``<w:evenAndOddHeaders/>`` into ``settings.xml`` if absent.
    Idempotent: a second call does not stack elements. Word will read
    even-page header/footer references from each section's
    ``<w:headerReference w:type="even">`` / ``<w:footerReference w:type="even">``
    children when this flag is set.

    Args:
        doc: The python-docx :class:`~docx.document.Document` whose
            settings part should carry the flag.

    Example:
        >>> from docx import Document
        >>> from docx_plus.layout import enable_distinct_even_odd_headers
        >>> doc = Document()
        >>> enable_distinct_even_odd_headers(doc)
    """
    settings = doc.settings.element
    if settings.find(qn("w:evenAndOddHeaders")) is not None:
        return
    new = el("w:evenAndOddHeaders")
    insert_before_first_anchor(
        settings, new, _EVEN_AND_ODD_HEADERS_LATER_SIBLINGS
    )


def disable_distinct_even_odd_headers(doc: DocxDocument) -> None:
    """Remove ``<w:evenAndOddHeaders/>`` from ``settings.xml`` if present.

    Idempotent: removing the flag when it is already absent is a no-op.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to mutate.
    """
    settings = doc.settings.element
    existing = settings.find(qn("w:evenAndOddHeaders"))
    if existing is not None:
        remove(existing)


__all__ = [
    "disable_distinct_even_odd_headers",
    "enable_distinct_even_odd_headers",
]
