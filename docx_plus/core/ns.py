"""OOXML namespace URIs and Clark-notation helper.

All XML element construction in the library uses these constants and the
:func:`qn` helper, so that a single change here propagates everywhere.
"""

from __future__ import annotations

from functools import cache

from docx_plus.core.errors import DocxPlusError

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
XML = "http://www.w3.org/XML/1998/namespace"

NSMAP: dict[str, str] = {
    "w": W,
    "w14": W14,
    "w15": W15,
    "w16cid": W16CID,
    "r": R,
    "mc": MC,
    "a": A,
    "xml": XML,
}

#: Prefixes declared on elements built in the main document namespaces.
#:
#: Deliberately narrower than :data:`NSMAP`, which is the *query* map and
#: has to know every prefix the library can match on. ``w15`` and
#: ``w16cid`` are only ever written into the comment side-parts
#: (``commentsExtended.xml``, ``people.xml``, ``commentsIds.xml``);
#: including them here would stamp a stray ``xmlns:w15`` onto every
#: element the library writes into ``document.xml``.
#: :func:`docx_plus.core.oxml.el` selects between this map and a
#: single-prefix one based on the element's own namespace.
BUILD_NSMAP: dict[str, str] = {
    "w": W,
    "w14": W14,
    "r": R,
    "mc": MC,
    "a": A,
    "xml": XML,
}


class InvalidNamespaceError(DocxPlusError, ValueError):
    """Raised by :func:`qn` for a malformed name or unknown prefix.

    Subclasses ``ValueError`` so existing ``except ValueError:`` clauses
    still catch it; also subclasses :class:`DocxPlusError` per SPEC §9.7.
    """


@cache
def qn(name: str) -> str:
    """Convert ``"prefix:local"`` to Clark notation ``"{namespace}local"``.

    Memoized, and unboundedly so: the input domain is the OOXML tag
    vocabulary, a few hundred literals fixed at authoring time. The
    resolver calls this from inside its hot loops — 563,000 times in a
    1000-paragraph sweep, 25% of the runtime — because writing ``qn("w:r")``
    at the callsite is what keeps that code legible. Caching buys the
    legibility back for free. Failures are not cached, so a bad name raises
    every time it is asked.

    Args:
        name: Qualified name in ``prefix:local`` form. ``prefix`` must be a
            key in :data:`NSMAP`.

    Returns:
        The Clark-notation form ``"{namespace-uri}local-name"`` used by lxml.

    Raises:
        InvalidNamespaceError: If ``name`` is not in ``prefix:local`` form,
            or the prefix is unknown.

    Example:
        >>> qn("w:tag")
        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tag'
    """
    if ":" not in name:
        raise InvalidNamespaceError(f"qn() expected 'prefix:local', got {name!r}")
    prefix, _, local = name.partition(":")
    try:
        uri = NSMAP[prefix]
    except KeyError as exc:
        raise InvalidNamespaceError(f"unknown namespace prefix {prefix!r} in {name!r}") from exc
    return f"{{{uri}}}{local}"


__all__ = [
    "A",
    "BUILD_NSMAP",
    "MC",
    "NSMAP",
    "R",
    "W",
    "W14",
    "W15",
    "W16CID",
    "XML",
    "InvalidNamespaceError",
    "qn",
]
