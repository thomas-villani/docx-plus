"""Document protection — SPEC §8.

Public API:

* :func:`protect_document` — enforce form-fill / read-only / comments-only mode
* :func:`unprotect_document` — remove protection
* :func:`is_protected` — predicate
"""

from __future__ import annotations

from docx_plus.protection.document import (
    ProtectionMode,
    is_protected,
    protect_document,
    unprotect_document,
)

__all__ = [
    "ProtectionMode",
    "is_protected",
    "protect_document",
    "unprotect_document",
]
