"""Read and modify content controls (SDTs) in an existing document.

The companion to :mod:`docx_plus.controls.builder`. Where ``builder`` writes
``w:sdt`` elements, this module discovers them, reports their values, sets new
values, or resets them to placeholder state.

The read side is intentionally schema-tolerant: it works on any document with
content controls, not just ones built by :class:`FormBuilder`. Type detection
dispatches on the marker child of ``w:sdtPr`` (``w:text``, ``w:dropDownList``,
``w:comboBox``, ``w:date``, ``w14:checkbox``, and the container/rich-text
markers Word writes but this module cannot set a value on).

Reading tolerates what Word actually emits, which is looser than what
:class:`FormBuilder` writes:

- ``w:tag`` is **optional and non-unique** in OOXML. Word's Developer-ribbon
  controls are written with ``<w:tag w:val=""/>`` unless the author types a
  tag, so empty and duplicate tags are the norm in real documents.
  :func:`list_controls` therefore keys nothing and returns document order;
  :func:`read_controls` is the keyed convenience built on top of it.
- A ``w:sdt`` with **no type marker at all** is a rich-text control (ECMA-376
  makes rich text the default), not an unrecognised element.
- Controls live in headers, footers, footnotes, and endnotes as well as the
  body, and every function here walks all of them.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import TYPE_CHECKING, Any, Literal

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from lxml import etree

from docx_plus.core import DocxPlusError
from docx_plus.core.ns import qn
from docx_plus.core.oxml import el, remove, sub, xpath

if TYPE_CHECKING:
    from docx.document import Document


ControlType = Literal[
    # Types whose value this module can read *and* write.
    "text",
    "dropdown",
    "combobox",
    "date",
    "checkbox",
    # Types this module reports but cannot set a value on. Rich text and the
    # container controls hold block-level content, not a single scalar.
    "richtext",
    "picture",
    "group",
    "repeating",
    "repeatingitem",
    "docpart",
    "citation",
    "bibliography",
    "equation",
]
ControlValueT = str | bool | datetime

#: Control types :func:`set_control_value` and :func:`clear_control` accept.
#: Everything else in :data:`ControlType` is read-only here.
WRITABLE_TYPES: frozenset[str] = frozenset({"text", "dropdown", "combobox", "date", "checkbox"})

#: ``w:sdtPr`` marker child -> control type, in probe order. ECMA-376 §17.5.2
#: makes these a choice group, so at most one is present; a ``w:sdt`` carrying
#: none of them is a rich-text control, which is why Word omits the marker on
#: the most common control it writes.
_TYPE_MARKERS: tuple[tuple[str, ControlType], ...] = (
    ("w:text", "text"),
    ("w:dropDownList", "dropdown"),
    ("w:comboBox", "combobox"),
    ("w:date", "date"),
    ("w14:checkbox", "checkbox"),
    ("w:picture", "picture"),
    ("w:group", "group"),
    ("w15:repeatingSection", "repeating"),
    ("w15:repeatingSectionItem", "repeatingitem"),
    ("w:docPartObj", "docpart"),
    ("w:docPartList", "docpart"),
    ("w:citation", "citation"),
    ("w:bibliography", "bibliography"),
    ("w:equation", "equation"),
    ("w:richText", "richtext"),
)

_PLACEHOLDER_STYLE_ID = "PlaceholderText"


# --------------------------------------------------------------------------
# Errors.
# --------------------------------------------------------------------------


class ControlNotFoundError(DocxPlusError, KeyError):
    """Raised when no content control with the requested tag exists.

    Subclasses ``KeyError`` so existing ``except KeyError:`` clauses still
    catch it; also subclasses :class:`DocxPlusError` per SPEC §9.7.
    """


class DuplicateTagError(DocxPlusError, ValueError):
    """Raised when a tag does not identify exactly one control.

    Raised by :func:`read_controls` when two controls share a non-empty key,
    and by :func:`set_control_value` / :func:`clear_control` when the requested
    tag matches more than one control — writing to an arbitrary one of them
    would silently corrupt the others' document.

    Controls whose ``w:tag`` is absent or empty are not a duplicate: they are
    unkeyable, so :func:`read_controls` omits them. Use :func:`list_controls`
    to see every control regardless of tag, and pass ``control_id`` to the
    writers to target one unambiguously.
    """


class ValueNotInListError(DocxPlusError, ValueError):
    """Raised by :func:`set_control_value` when a dropdown value has no match."""


class ControlTypeError(DocxPlusError, TypeError):
    """Raised when a value's Python type does not match the control's type."""


# --------------------------------------------------------------------------
# ControlValue.
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class ControlValue:
    """A single content control's identity, type, and current value.

    Attributes:
        tag: The control's ``w:tag`` value, or ``None`` when the control has no
            ``w:tag`` element at all. An empty string means the element is
            present but its ``w:val`` is empty — the shape Word writes for a
            control the author never tagged. Neither is unique, so a tag is a
            *label*, not a primary key; use ``control_id`` for identity.
        alias: The control's ``w:alias`` value (UI label), or ``None``. Also
            not unique — aliases are human labels and repeat freely.
        control_type: One of the values in :data:`ControlType`.
        value: The current value:

            - text/dropdown/combobox: ``str`` if filled, ``None`` if showing
              placeholder.
            - date: :class:`~datetime.datetime` if filled, ``None`` otherwise.
            - checkbox: always ``bool`` (no placeholder concept).
            - every other type: the control's concatenated text, which for a
              container control is the text of everything it wraps.

        is_placeholder: True if the control is showing its placeholder text
            (``w:showingPlcHdr`` present in sdtPr). Always ``False`` for
            checkboxes.
        control_id: The control's ``w:id`` value, or ``None`` if absent or
            non-numeric. This is OOXML's actual identity field; pass it to
            :func:`set_control_value` to disambiguate a repeated tag. Word
            does not guarantee uniqueness across a document merge, so treat a
            collision here as possible but rare.
        index: Zero-based position in :func:`list_controls` order — document
            order within :attr:`location`, parts visited body-first. Stable for
            a given document, *not* stable across edits that add controls.
        location: Which story the control lives in: ``"body"``, ``"footnotes"``,
            ``"endnotes"``, or ``"header:S:WHICH"`` / ``"footer:S:WHICH"``
            where ``S`` is the 1-based section index and ``WHICH`` is
            ``primary``, ``first``, or ``even``.
    """

    tag: str | None
    alias: str | None
    control_type: ControlType
    value: ControlValueT | None
    is_placeholder: bool
    control_id: int | None = None
    index: int = 0
    location: str = "body"


# --------------------------------------------------------------------------
# Public API.
# --------------------------------------------------------------------------


def list_controls(doc: Document) -> list[ControlValue]:
    """Return every content control in ``doc``, in document order.

    The unkeyed primitive that :func:`read_controls` is built on. Nothing is
    dropped and nothing can collide, so this is the function to reach for on
    documents Word produced rather than :class:`FormBuilder`: it reports
    controls with an empty tag, with no tag element, of rich-text or container
    type, and in headers, footers, footnotes, and endnotes.

    Args:
        doc: The python-docx Document to inspect.

    Returns:
        A list of :class:`ControlValue`, one per ``w:sdt``, ordered by story
        (body first, then headers/footers by section, then notes) and by
        document order within each story. Each entry's ``index`` matches its
        position in this list.
    """
    out: list[ControlValue] = []
    for sdt, location in _iter_sdts(doc):
        info = _read_sdt(sdt, index=len(out), location=location)
        if info is None:
            continue
        out.append(info)
    return out


def read_controls(
    doc: Document,
    *,
    by: Literal["tag", "alias"] = "tag",
) -> dict[str, ControlValue]:
    """Return the content controls in ``doc`` that have a usable key.

    A convenience wrapper over :func:`list_controls` for the common case of a
    form whose controls carry deliberate, distinct tags — the shape
    :class:`FormBuilder` writes.

    **Controls without a usable key are omitted.** In OOXML ``w:tag`` is
    neither required nor unique, and Word writes ``<w:tag w:val=""/>`` for any
    control the author did not explicitly tag, so on a real Word document this
    can omit most of them. Use :func:`list_controls` when you need every
    control, or ``by="alias"`` when the document labels controls rather than
    tagging them.

    Args:
        doc: The python-docx Document to inspect.
        by: Either ``"tag"`` (default) — key on ``w:tag`` — or ``"alias"`` —
            key on ``w:alias``. Either way, controls whose key is absent or
            empty are skipped.

    Returns:
        Mapping from key to :class:`ControlValue`.

    Raises:
        DuplicateTagError: If two controls share the same non-empty key. This
            is genuine ambiguity, unlike an absent key, so it is reported
            rather than silently resolved.
    """
    out: dict[str, ControlValue] = {}
    for info in list_controls(doc):
        key = info.tag if by == "tag" else info.alias
        if not key:
            continue
        if key in out:
            raise DuplicateTagError(
                f"duplicate {by} {key!r} encountered while reading controls; "
                f"use list_controls() to read a document with repeated {by}s",
            )
        out[key] = info
    return out


def set_control_value(
    doc: Document,
    tag: str | None,
    value: ControlValueT,
    *,
    control_id: int | None = None,
) -> None:
    """Set the value of a control identified by ``tag`` (or ``control_id``).

    Args:
        doc: The python-docx Document to modify.
        tag: The control's ``w:tag`` value. May be ``None`` when ``control_id``
            is given.
        value: The new value. Type must match the control type:

            - text: ``str``
            - dropdown / combobox: ``str``
            - date: :class:`~datetime.datetime`
            - checkbox: ``bool``

        control_id: The control's ``w:id`` value, from
            :attr:`ControlValue.control_id`. Selects one control directly and
            ignores ``tag`` — the way to write to a control whose tag is
            empty, absent, or shared with others.

    Raises:
        ControlNotFoundError: If no control matches.
        DuplicateTagError: If ``tag`` matches more than one control and no
            ``control_id`` was given to disambiguate.
        ControlTypeError: If ``value``'s type does not match the control type,
            or the control is one of the rich-text/container types this
            module cannot set a scalar value on.
        ValueNotInListError: For a dropdown when ``value`` matches neither
            ``w:value`` nor ``w:displayText`` of any list item.
    """
    sdt = _select_sdt(doc, tag, control_id)
    sdt_pr = _sdt_pr(sdt)
    sdt_content = _sdt_content(sdt)
    control_type = _require_writable(sdt, tag, control_id)

    if control_type == "checkbox":
        if not isinstance(value, bool):
            raise ControlTypeError(
                f"checkbox control {tag!r} requires bool; got {type(value).__name__}",
            )
        _set_checkbox(sdt_pr, sdt_content, checked=value)
        return

    if control_type == "date":
        if not isinstance(value, datetime):
            raise ControlTypeError(
                f"date control {tag!r} requires datetime; got {type(value).__name__}",
            )
        _set_date(sdt_pr, sdt_content, value)
        _clear_placeholder_flag(sdt_pr)
        return

    # text / dropdown / combobox
    if not isinstance(value, str):
        raise ControlTypeError(
            f"{control_type} control {tag!r} requires str; got {type(value).__name__}",
        )

    if control_type == "text":
        _replace_sdt_content_text(sdt_content, value)
    elif control_type == "dropdown":
        display = _resolve_dropdown_value(sdt_pr, value, allow_freeform=False, tag=tag)
        _replace_sdt_content_text(sdt_content, display)
    else:  # combobox
        display = _resolve_dropdown_value(sdt_pr, value, allow_freeform=True, tag=tag)
        _replace_sdt_content_text(sdt_content, display)

    _clear_placeholder_flag(sdt_pr)


def clear_control(
    doc: Document,
    tag: str | None,
    *,
    control_id: int | None = None,
) -> None:
    """Reset a control to its placeholder state.

    For text/dropdown/combobox/date: re-adds ``w:showingPlcHdr`` to sdtPr and
    re-applies the ``PlaceholderText`` rStyle to every run in sdtContent. The
    placeholder text itself is preserved in place (whatever sdtContent
    currently holds).

    For checkbox: resets the checked flag to ``0`` and the glyph to
    ``☐``. Checkboxes have no placeholder mode.

    Args:
        doc: The python-docx Document to modify.
        tag: The control's ``w:tag`` value. May be ``None`` when ``control_id``
            is given.
        control_id: The control's ``w:id`` value; selects one control directly
            and ignores ``tag``. See :func:`set_control_value`.

    Raises:
        ControlNotFoundError: If no control matches.
        DuplicateTagError: If ``tag`` matches more than one control and no
            ``control_id`` was given to disambiguate.
        ControlTypeError: If the control is a rich-text or container type,
            which has no placeholder state to reset.
    """
    sdt = _select_sdt(doc, tag, control_id)
    sdt_pr = _sdt_pr(sdt)
    sdt_content = _sdt_content(sdt)
    control_type = _require_writable(sdt, tag, control_id)

    if control_type == "checkbox":
        _set_checkbox(sdt_pr, sdt_content, checked=False)
        return

    _set_placeholder_flag(sdt_pr)
    for run in sdt_content.findall(qn("w:r")):
        rpr = run.find(qn("w:rPr"))
        if rpr is None:
            rpr = el("w:rPr")
            run.insert(0, rpr)
        for existing in rpr.findall(qn("w:rStyle")):
            remove(existing)
        rstyle = el("w:rStyle", **{"w:val": _PLACEHOLDER_STYLE_ID})
        rpr.insert(0, rstyle)


# --------------------------------------------------------------------------
# Shared SDT classification (also used by _testing/ooxml_asserts).
# --------------------------------------------------------------------------


def _classify_sdt(sdt: etree._Element) -> ControlType | None:
    """Return the control type for an SDT, or ``None`` if it has no ``w:sdtPr``.

    A ``w:sdt`` carrying none of the markers in :data:`_TYPE_MARKERS` is a
    rich-text control, not an unrecognised one — ECMA-376 §17.5.2 makes rich
    text the default, and Word relies on that by omitting the marker. Only a
    structurally broken SDT (no ``w:sdtPr`` at all) returns ``None``.
    """
    sdt_pr = sdt.find(qn("w:sdtPr"))
    if sdt_pr is None:
        return None
    for marker, kind in _TYPE_MARKERS:
        if sdt_pr.find(qn(marker)) is not None:
            return kind
    return "richtext"


# --------------------------------------------------------------------------
# Internals.
# --------------------------------------------------------------------------


#: Header/footer accessors on a :class:`docx.section.Section`, paired with the
#: ``WHICH`` component of :attr:`ControlValue.location`.
_HDRFTR_SLOTS: tuple[tuple[str, str, str], ...] = (
    ("header", "primary", "header"),
    ("header", "first", "first_page_header"),
    ("header", "even", "even_page_header"),
    ("footer", "primary", "footer"),
    ("footer", "first", "first_page_footer"),
    ("footer", "even", "even_page_footer"),
)


def _iter_story_roots(doc: Document) -> list[tuple[str, etree._Element]]:
    """Return ``(location, root)`` for every story that can hold a ``w:sdt``.

    Body first, then each section's header/footer definitions, then the
    footnote and endnote parts. Strictly non-mutating: a header/footer is only
    touched when ``is_linked_to_previous`` is ``False``, because python-docx
    *creates* the part on first access to an undefined one, and an inherited
    definition would otherwise be walked once per section that inherits it.
    """
    roots: list[tuple[str, etree._Element]] = [("body", doc.element.body)]

    for section_index, section in enumerate(doc.sections, start=1):
        for kind, which, attr in _HDRFTR_SLOTS:
            hdrftr: Any = getattr(section, attr)
            if hdrftr.is_linked_to_previous:
                continue
            roots.append((f"{kind}:{section_index}:{which}", hdrftr._element))

    for location, rel_type in (("footnotes", RT.FOOTNOTES), ("endnotes", RT.ENDNOTES)):
        try:
            part: Any = doc.part.part_related_by(rel_type)
        except KeyError:
            continue
        element = getattr(part, "element", None)
        if isinstance(element, etree._Element):
            roots.append((location, element))

    return roots


def _iter_sdts(doc: Document) -> list[tuple[etree._Element, str]]:
    """Return ``(sdt, location)`` for every content control in the document."""
    out: list[tuple[etree._Element, str]] = []
    for location, root in _iter_story_roots(doc):
        out.extend((s, location) for s in xpath(root, ".//w:sdt") if isinstance(s, etree._Element))
    return out


def _select_sdt(
    doc: Document,
    tag: str | None,
    control_id: int | None,
) -> etree._Element:
    """Return the one SDT identified by ``control_id``, else by ``tag``.

    Raises:
        ControlNotFoundError: If nothing matches, or neither selector is given.
        DuplicateTagError: If ``tag`` matches more than one control. Writing to
            an arbitrary match would leave the others silently untouched, so
            ambiguity is refused rather than resolved.
    """
    if control_id is not None:
        for sdt, _location in _iter_sdts(doc):
            if _sdt_id(sdt) == control_id:
                return sdt
        raise ControlNotFoundError(f"no content control with id {control_id!r}")

    if tag is None:
        raise ControlNotFoundError("one of tag or control_id is required")

    matches = [sdt for sdt, _location in _iter_sdts(doc) if _sdt_tag(sdt) == tag]
    if not matches:
        raise ControlNotFoundError(f"no content control with tag {tag!r}")
    if len(matches) > 1:
        ids = [_sdt_id(sdt) for sdt in matches]
        raise DuplicateTagError(
            f"tag {tag!r} matches {len(matches)} controls (ids {ids!r}); "
            f"pass control_id= to choose one",
        )
    return matches[0]


def _sdt_tag(sdt: etree._Element) -> str | None:
    """Return the SDT's ``w:tag`` value, or ``None`` if it has no tag element."""
    sdt_pr = sdt.find(qn("w:sdtPr"))
    if sdt_pr is None:
        return None
    tag_el = sdt_pr.find(qn("w:tag"))
    if tag_el is None:
        return None
    return tag_el.get(qn("w:val")) or ""


def _sdt_id(sdt: etree._Element) -> int | None:
    """Return the SDT's ``w:id`` value, or ``None`` if absent or non-numeric."""
    sdt_pr = sdt.find(qn("w:sdtPr"))
    if sdt_pr is None:
        return None
    id_el = sdt_pr.find(qn("w:id"))
    if id_el is None:
        return None
    raw = id_el.get(qn("w:val"))
    if raw is None:
        return None
    try:
        return int(raw)
    except ValueError:
        return None


def _require_writable(
    sdt: etree._Element,
    tag: str | None,
    control_id: int | None,
) -> ControlType:
    """Return the control's type, rejecting the ones with no scalar value."""
    control_type = _classify_sdt(sdt)
    label = f"tag {tag!r}" if control_id is None else f"id {control_id!r}"
    if control_type is None:  # pragma: no cover - _sdt_pr already rejected this
        raise ControlNotFoundError(f"control with {label} is malformed: missing w:sdtPr")
    if control_type not in WRITABLE_TYPES:
        raise ControlTypeError(
            f"control with {label} is a {control_type} control, which holds "
            f"block-level content rather than a single value; "
            f"writable types are {sorted(WRITABLE_TYPES)}",
        )
    return control_type


def _sdt_pr(sdt: etree._Element) -> etree._Element:
    sdt_pr = sdt.find(qn("w:sdtPr"))
    if sdt_pr is None:
        raise ControlNotFoundError("malformed SDT: missing w:sdtPr")
    return sdt_pr


def _sdt_content(sdt: etree._Element) -> etree._Element:
    sdt_content = sdt.find(qn("w:sdtContent"))
    if sdt_content is None:
        raise ControlNotFoundError("malformed SDT: missing w:sdtContent")
    return sdt_content


def _read_sdt(
    sdt: etree._Element,
    *,
    index: int = 0,
    location: str = "body",
) -> ControlValue | None:
    """Read one ``w:sdt``. ``None`` only for a structurally broken element.

    Every well-formed SDT is reportable: an absent ``w:tag`` yields
    ``tag=None`` and an absent type marker yields ``richtext``, because
    dropping either would silently hide controls Word routinely writes.
    """
    sdt_pr = sdt.find(qn("w:sdtPr"))
    if sdt_pr is None:
        return None
    control_type = _classify_sdt(sdt)
    if control_type is None:  # pragma: no cover - implied by the sdt_pr check
        return None

    tag = _sdt_tag(sdt)

    alias_el = sdt_pr.find(qn("w:alias"))
    alias = alias_el.get(qn("w:val")) if alias_el is not None else None

    is_placeholder = sdt_pr.find(qn("w:showingPlcHdr")) is not None
    sdt_content = sdt.find(qn("w:sdtContent"))

    value: ControlValueT | None
    if control_type == "checkbox":
        value = _read_checkbox_value(sdt_pr)
        # Checkboxes never carry placeholder semantics — Word always renders
        # checked-or-unchecked.
        is_placeholder = False
    elif control_type == "date":
        value = _read_date_value(sdt_pr)
        if is_placeholder:
            value = None
    else:
        text = _collect_text(sdt_content) if sdt_content is not None else ""
        value = None if is_placeholder else text

    return ControlValue(
        tag=tag,
        alias=alias,
        control_type=control_type,
        value=value,
        is_placeholder=is_placeholder,
        control_id=_sdt_id(sdt),
        index=index,
        location=location,
    )


def _read_checkbox_value(sdt_pr: etree._Element) -> bool:
    checkbox = sdt_pr.find(qn("w14:checkbox"))
    if checkbox is None:
        return False
    checked = checkbox.find(qn("w14:checked"))
    if checked is None:
        return False
    raw = checked.get(qn("w14:val"))
    return raw not in (None, "0", "false")


def _read_date_value(sdt_pr: etree._Element) -> datetime | None:
    date_el = sdt_pr.find(qn("w:date"))
    if date_el is None:
        return None
    iso = date_el.get(qn("w:fullDate"))
    if iso is None:
        return None
    # The ``Z`` → ``+00:00`` swap supports Python 3.10's ``fromisoformat``,
    # which doesn't accept ``Z``. Drop the replacement once the minimum
    # Python bumps to 3.11+, which parses ``Z`` natively.
    try:
        return datetime.fromisoformat(iso.replace("Z", "+00:00"))
    except ValueError:
        return None


def _collect_text(sdt_content: etree._Element) -> str:
    parts: list[str] = []
    for t in sdt_content.iter(qn("w:t")):
        if t.text is not None:
            parts.append(t.text)
    return "".join(parts)


def _resolve_dropdown_value(
    sdt_pr: etree._Element,
    value: str,
    *,
    allow_freeform: bool,
    tag: str | None,
) -> str:
    """Match ``value`` against listItem entries; return the displayText to render.

    Match priority: ``w:value`` first, then ``w:displayText``. If no match and
    ``allow_freeform`` is ``True`` (combobox), return ``value`` verbatim;
    otherwise raise :class:`ValueNotInListError`. The auto-prepended empty-value
    placeholder list-item is skipped during matching so it cannot shadow real
    entries.
    """
    list_el = sdt_pr.find(qn("w:dropDownList"))
    if list_el is None:
        list_el = sdt_pr.find(qn("w:comboBox"))
    if list_el is None:
        raise ControlNotFoundError(
            f"control with tag {tag!r} has no list element",
        )

    items = list_el.findall(qn("w:listItem"))
    for item in items:
        if item.get(qn("w:value")) == "":
            continue
        if item.get(qn("w:value")) == value:
            return item.get(qn("w:displayText")) or value
    for item in items:
        if item.get(qn("w:value")) == "":
            continue
        if item.get(qn("w:displayText")) == value:
            return item.get(qn("w:displayText")) or value

    if allow_freeform:
        return value
    raise ValueNotInListError(
        f"dropdown {tag!r} has no list item matching {value!r}",
    )


def _replace_sdt_content_text(sdt_content: etree._Element, text: str) -> None:
    """Replace sdtContent's children with a single plain run containing ``text``."""
    for child in list(sdt_content):
        sdt_content.remove(child)
    run = el("w:r")
    text_el = sub(run, "w:t")
    text_el.text = text
    if text != text.strip() or "\n" in text:
        text_el.set(qn("xml:space"), "preserve")
    sdt_content.append(run)


def _clear_placeholder_flag(sdt_pr: etree._Element) -> None:
    flag = sdt_pr.find(qn("w:showingPlcHdr"))
    if flag is not None:
        remove(flag)


def _set_placeholder_flag(sdt_pr: etree._Element) -> None:
    """Add ``w:showingPlcHdr`` to sdtPr in the schema-correct position.

    Schema order observed in builder: ``[alias?], tag, id, [showingPlcHdr],
    <type-marker>``. We re-insert immediately before the type marker child so
    re-clearing a previously-cleared control restores the canonical order.
    """
    existing = sdt_pr.find(qn("w:showingPlcHdr"))
    if existing is not None:
        return
    flag = el("w:showingPlcHdr")
    type_marker_tags = (
        qn("w:text"),
        qn("w:dropDownList"),
        qn("w:comboBox"),
        qn("w:date"),
        qn("w14:checkbox"),
    )
    for child in sdt_pr:
        if child.tag in type_marker_tags:
            child.addprevious(flag)
            return
    sdt_pr.append(flag)


def _set_date(
    sdt_pr: etree._Element,
    sdt_content: etree._Element,
    value: datetime,
) -> None:
    date_el = sdt_pr.find(qn("w:date"))
    if date_el is None:
        date_el = sub(sdt_pr, "w:date")
    date_el.set(qn("w:fullDate"), value.isoformat())
    fmt_el = date_el.find(qn("w:dateFormat"))
    fmt_string = fmt_el.get(qn("w:val")) if fmt_el is not None else None
    rendered = _render_date(value, fmt_string)
    _replace_sdt_content_text(sdt_content, rendered)


def _render_date(value: datetime, fmt: str | None) -> str:
    """Render ``value`` for display in sdtContent.

    Word's date-format tokens (``M/d/yyyy``, ``dddd, MMMM d, yyyy``) only
    overlap partially with Python's strftime tokens. Translating them in full
    is out of scope for v0.1 — the canonical machine value lives in
    ``w:date/@w:fullDate`` (ISO 8601), so the rendered text only needs to be
    a sane human-readable form. We special-case the common Word default and
    fall back to ISO date for anything else.

    Word re-renders the displayed text on next open from the canonical
    ``w:fullDate`` using the saved ``w:dateFormat``, so a "looks ISO until
    Word touches it" output is correct, just not visually identical to
    what Word would produce in the meantime.
    """
    if fmt == "M/d/yyyy" or fmt is None:
        return f"{value.month}/{value.day}/{value.year}"
    return value.date().isoformat()


def _set_checkbox(
    sdt_pr: etree._Element,
    sdt_content: etree._Element,
    *,
    checked: bool,
) -> None:
    checkbox = sdt_pr.find(qn("w14:checkbox"))
    if checkbox is None:
        raise ControlNotFoundError("malformed checkbox SDT: missing w14:checkbox")
    checked_el = checkbox.find(qn("w14:checked"))
    if checked_el is None:
        checked_el = sub(checkbox, "w14:checked")
    checked_el.set(qn("w14:val"), "1" if checked else "0")

    glyph = "☒" if checked else "☐"
    for run in sdt_content.findall(qn("w:r")):
        for t in run.findall(qn("w:t")):
            t.text = glyph
            return
    # No existing run — synthesize a minimal one.
    run = sub(sdt_content, "w:r")
    text_el = sub(run, "w:t")
    text_el.text = glyph


__all__ = [
    "ControlNotFoundError",
    "ControlType",
    "ControlTypeError",
    "WRITABLE_TYPES",
    "ControlValue",
    "DuplicateTagError",
    "ValueNotInListError",
    "clear_control",
    "list_controls",
    "read_controls",
    "set_control_value",
]
