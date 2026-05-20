"""Library base error class, isolated to break an import cycle.

:class:`DocxPlusError` lives here — in a module that imports nothing else
from ``docx_plus.core`` — so the other ``core`` submodules (``ids``,
``ns``, …) can subclass it with a plain top-of-file import instead of the
``# noqa: E402`` ordering dance that an in-``__init__`` definition would
force. ``core/__init__`` re-exports it, so the documented short form
``from docx_plus.core import DocxPlusError`` is unchanged. SPEC §9.7.
"""

from __future__ import annotations


class DocxPlusError(Exception):
    """Base class for all library-raised errors.

    Every typed error in docx_plus subclasses this so callers can catch the
    library's failures without catching unrelated ``ValueError``/``RuntimeError``
    instances from python-docx or lxml.
    """


__all__ = ["DocxPlusError"]
