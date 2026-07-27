"""Reading list definitions back out of ``numbering.xml``.

The read side of :mod:`docx_plus.numbering.define`. It reports what is
actually in the part — including definitions Word or another tool wrote,
and including the nine ``abstractNum`` entries python-docx's bundled
template ships in every fresh document.

Like every reader in the library this never fabricates a part: a
document with no ``numbering.xml`` reads as an empty list rather than
gaining one as a side effect of being inspected.

This module imports only from ``docx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, cast

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import XmlPart

from docx_plus.core.ns import qn
from docx_plus.core.oxml import xpath

if TYPE_CHECKING:
    from docx.document import Document
    from lxml import etree


@dataclass(frozen=True)
class ListLevel:
    """One outline level of a definition, as found in the document.

    The read-side counterpart of
    :class:`~docx_plus.numbering.LevelDefinition`. Every field mirrors an
    optional child of ``<w:lvl>``, so ``None`` means "the element is
    absent" — which Word reads as its own default, not as zero.

    Attributes:
        level: Zero-based outline depth (``w:ilvl``).
        fmt: ``w:numFmt`` value, e.g. ``"decimal"`` or ``"bullet"``.
        text: ``w:lvlText`` pattern or literal bullet glyph.
        start: ``w:start`` value.
        indent: Left indent in twips from the level's ``w:ind``.
        hanging: Hanging indent in twips from the same.
        justify: ``w:lvlJc`` value.
        suffix: ``w:suff`` value. ``None`` means the element is absent,
            which Word treats as ``"tab"``.
        restart_after: ``w:lvlRestart`` value.
        font: ``w:ascii`` from the level's ``w:rPr/w:rFonts``.
    """

    level: int
    fmt: str | None = None
    text: str | None = None
    start: int | None = None
    indent: int | None = None
    hanging: int | None = None
    justify: str | None = None
    suffix: str | None = None
    restart_after: int | None = None
    font: str | None = None


@dataclass(frozen=True)
class ListDefinition:
    """A ``<w:num>`` instance together with the definition behind it.

    Attributes:
        num_id: The ``w:numId`` paragraphs reference.
        abstract_id: The ``w:abstractNumId`` it points at. ``None`` if
            the ``w:num`` carries no reference — malformed, but present
            in the wild.
        levels: The abstract definition's levels, outermost first. Empty
            if the reference is dangling.
        name: The definition's ``w:name``, if any.
        style_link: ``w:styleLink`` — the style this definition is the
            numbering for.
        num_style_link: ``w:numStyleLink`` — the style whose numbering
            this definition defers to.
        multi_level_type: ``w:multiLevelType``.
        start_overrides: ``{level: start}`` for every
            ``w:lvlOverride/w:startOverride`` on the instance. This is
            what distinguishes a restarted sequence from the original —
            see :func:`~docx_plus.numbering.restart_list`.
    """

    num_id: int
    abstract_id: int | None
    levels: tuple[ListLevel, ...]
    name: str | None = None
    style_link: str | None = None
    num_style_link: str | None = None
    multi_level_type: str | None = None
    start_overrides: tuple[tuple[int, int], ...] = ()


def read_list_definitions(doc: Document) -> list[ListDefinition]:
    """Return every list definition in ``doc``, in ``numbering.xml`` order.

    Note:
        A fresh ``Document()`` is **not** empty here. python-docx's
        bundled template ships nine ``abstractNum`` entries and nine
        ``num`` instances covering the built-in ``List Bullet`` and
        ``List Number`` styles, so a document you have not touched
        already reports nine definitions.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to scan.

    Returns:
        One :class:`ListDefinition` per ``<w:num>``. Returns ``[]`` if
        the document has no ``numbering.xml``.

    Example:
        >>> from docx import Document
        >>> from docx_plus.numbering import define_bullet_list, read_list_definitions
        >>> doc = Document()
        >>> num = define_bullet_list(doc)
        >>> mine = [d for d in read_list_definitions(doc) if d.num_id == num]
        >>> mine[0].levels[0].fmt
        'bullet'
    """
    root = _numbering_root(doc)
    if root is None:
        return []

    abstracts = {
        abstract.get(qn("w:abstractNumId")): abstract for abstract in xpath(root, "./w:abstractNum")
    }

    definitions: list[ListDefinition] = []
    for num in xpath(root, "./w:num"):
        num_id = _int_attr(num, "w:numId")
        if num_id is None:
            continue  # a w:num with no id cannot be referenced; skip it
        raw_abstract = _child_val(num, "w:abstractNumId")
        abstract = abstracts.get(raw_abstract) if raw_abstract is not None else None
        definitions.append(
            ListDefinition(
                num_id=num_id,
                abstract_id=_as_int(raw_abstract),
                levels=_read_levels(abstract),
                name=_child_val(abstract, "w:name"),
                style_link=_child_val(abstract, "w:styleLink"),
                num_style_link=_child_val(abstract, "w:numStyleLink"),
                multi_level_type=_child_val(abstract, "w:multiLevelType"),
                start_overrides=_read_start_overrides(num),
            )
        )
    return definitions


# ---------------------------------------------------------------------------
# Internals.
# ---------------------------------------------------------------------------


def _numbering_root(doc: Document) -> etree._Element | None:
    """Return the ``w:numbering`` root, or ``None`` if the part is absent.

    Reads the relationship directly rather than through
    ``doc.part.numbering_part``, which fabricates a missing part via an
    unimplemented stub and raises ``NotImplementedError``.
    """
    try:
        part = cast("XmlPart", doc.part.part_related_by(RT.NUMBERING))
    except KeyError:
        return None
    return cast("etree._Element", part.element)


def _read_levels(abstract: etree._Element | None) -> tuple[ListLevel, ...]:
    if abstract is None:
        return ()
    levels = []
    for lvl in xpath(abstract, "./w:lvl"):
        ilvl = _int_attr(lvl, "w:ilvl")
        if ilvl is None:
            continue
        ind = lvl.find(f"{qn('w:pPr')}/{qn('w:ind')}")
        fonts = lvl.find(f"{qn('w:rPr')}/{qn('w:rFonts')}")
        levels.append(
            ListLevel(
                level=ilvl,
                fmt=_child_val(lvl, "w:numFmt"),
                text=_child_val(lvl, "w:lvlText"),
                start=_as_int(_child_val(lvl, "w:start")),
                indent=_as_int(ind.get(qn("w:left")) if ind is not None else None),
                hanging=_as_int(ind.get(qn("w:hanging")) if ind is not None else None),
                justify=_child_val(lvl, "w:lvlJc"),
                suffix=_child_val(lvl, "w:suff"),
                restart_after=_as_int(_child_val(lvl, "w:lvlRestart")),
                font=fonts.get(qn("w:ascii")) if fonts is not None else None,
            )
        )
    return tuple(levels)


def _read_start_overrides(num: etree._Element) -> tuple[tuple[int, int], ...]:
    overrides = []
    for override in xpath(num, "./w:lvlOverride"):
        ilvl = _int_attr(override, "w:ilvl")
        start = _as_int(_child_val(override, "w:startOverride"))
        if ilvl is not None and start is not None:
            overrides.append((ilvl, start))
    return tuple(overrides)


def _child_val(parent: etree._Element | None, tag: str) -> str | None:
    """Return a child element's ``w:val``, or ``None`` if either is absent."""
    if parent is None:
        return None
    child = parent.find(qn(tag))
    if child is None:
        return None
    value = child.get(qn("w:val"))
    return str(value) if value is not None else None


def _int_attr(elem: etree._Element, attr: str) -> int | None:
    return _as_int(elem.get(qn(attr)))


def _as_int(raw: str | None) -> int | None:
    """Parse an OOXML integer attribute, tolerating garbage.

    Malformed ids are common enough in documents produced by other tools
    that raising here would make the reader useless for exactly the files
    worth inspecting.
    """
    if raw is None:
        return None
    try:
        return int(raw)
    except ValueError:
        return None


__all__ = ["ListDefinition", "ListLevel", "read_list_definitions"]
