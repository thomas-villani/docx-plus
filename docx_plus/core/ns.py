"""OOXML namespace URIs and Clark-notation helper.

All XML element construction in the library uses these constants and the
:func:`qn` helper, so that a single change here propagates everywhere.
"""

from __future__ import annotations

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
XML = "http://www.w3.org/XML/1998/namespace"

NSMAP: dict[str, str] = {
    "w": W,
    "w14": W14,
    "r": R,
    "mc": MC,
    "a": A,
    "xml": XML,
}


def qn(name: str) -> str:
    """Convert ``"prefix:local"`` to Clark notation ``"{namespace}local"``.

    Args:
        name: Qualified name in ``prefix:local`` form. ``prefix`` must be a
            key in :data:`NSMAP`.

    Returns:
        The Clark-notation form ``"{namespace-uri}local-name"`` used by lxml.

    Raises:
        ValueError: If ``name`` is not in ``prefix:local`` form, or the prefix
            is unknown.

    Example:
        >>> qn("w:tag")
        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tag'
    """
    if ":" not in name:
        raise ValueError(f"qn() expected 'prefix:local', got {name!r}")
    prefix, _, local = name.partition(":")
    try:
        uri = NSMAP[prefix]
    except KeyError as exc:
        raise ValueError(f"unknown namespace prefix {prefix!r} in {name!r}") from exc
    return f"{{{uri}}}{local}"


__all__ = ["A", "MC", "NSMAP", "R", "W", "W14", "XML", "qn"]
