"""Core foundation: namespaces, XML helpers, ID registry, package parts.

This subpackage is the only dependency target permitted to capability modules
(``styles/``, ``controls/``, ``fields/``, ``protection/``). See SPEC §9.1.
"""


class DocxPlusError(Exception):
    """Base class for all library-raised errors.

    Every typed error in docx_plus subclasses this so callers can catch the
    library's failures without catching unrelated ``ValueError``/``RuntimeError``
    instances from python-docx or lxml.
    """


__all__ = ["DocxPlusError"]
