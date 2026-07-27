"""Custom numbering and list definitions — v0.5.

The largest remaining ``python-docx`` gap. python-docx exposes a
``NumberingPart`` and ``len()`` of its definitions, and that is all:
``docx/oxml/numbering.py`` has no ``CT_AbstractNum`` and no ``CT_Lvl``,
so nothing in it can express what a list looks like. Building a
multilevel list means hand-writing XML — and the part that would create
``numbering.xml`` when it is missing raises ``NotImplementedError``.

OOXML splits a list in two. A ``<w:abstractNum>`` is the definition — up
to nine ``<w:lvl>`` children giving each depth its number format, level
text, start value, indents, and bullet font. A ``<w:num>`` is an
*instance* of one, and a paragraph's ``<w:numPr>`` references the
instance. That indirection is the whole trick behind restarting: two
instances of one definition are independent counters that look identical.

Public surface:

- :class:`LevelDefinition` — one outline level
- :func:`define_list_definition` — the primitive
- :func:`define_bullet_list` / :func:`define_numbered_list` — presets
  using Word's own glyph and format cycles
- :func:`apply_list` / :func:`remove_list` — paragraph membership
- :func:`restart_list` — begin a fresh sequence over the same definition
- :func:`read_list_definitions` — read back, returning
  :class:`ListDefinition` / :class:`ListLevel`
- :class:`NumIdRegistry` / :class:`AbstractNumIdRegistry` — share
  allocators across an editing session
- :class:`InvalidLevelError`, :class:`ListDefinitionNotFoundError`

Not covered: linking a numbering definition into a *style* definition
(``w:style/w:pPr/w:numPr``). ``styles.modify`` already owns writing into
``w:style`` and carries the schema orders for it, so that belongs there;
see ``ROADMAP.md``. Until then a style like ``ListBullet`` created by
:func:`~docx_plus.styles.ensure_style` carries no numbering of its own —
apply a definition to the paragraphs directly.

See ``ROADMAP.md`` for where this capability was scoped.
"""

from __future__ import annotations

from docx_plus.numbering.apply import (
    ListDefinitionNotFoundError,
    apply_list,
    remove_list,
    restart_list,
)
from docx_plus.numbering.define import (
    MAX_LEVELS,
    InvalidLevelError,
    Justification,
    LevelDefinition,
    MultiLevelType,
    Suffix,
    define_bullet_list,
    define_list_definition,
    define_numbered_list,
)
from docx_plus.numbering.read import ListDefinition, ListLevel, read_list_definitions
from docx_plus.numbering.registry import AbstractNumIdRegistry, NumIdRegistry

__all__ = [
    "MAX_LEVELS",
    "AbstractNumIdRegistry",
    "InvalidLevelError",
    "Justification",
    "LevelDefinition",
    "ListDefinition",
    "ListDefinitionNotFoundError",
    "ListLevel",
    "MultiLevelType",
    "NumIdRegistry",
    "Suffix",
    "apply_list",
    "define_bullet_list",
    "define_list_definition",
    "define_numbered_list",
    "read_list_definitions",
    "remove_list",
    "restart_list",
]
