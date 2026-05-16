"""Read and modify content controls (SDTs) in an existing document.

The companion to :mod:`docx_plus.controls.builder`. Where ``builder`` writes
``w:sdt`` elements, this module discovers them, reports their values, sets new
values, or resets them to placeholder state.

The read side is intentionally schema-tolerant: it works on any document with
content controls, not just ones built by :class:`FormBuilder`. Type detection
dispatches on the marker child of ``w:sdtPr`` (``w:text``, ``w:dropDownList``,
``w:comboBox``, ``w:date``, ``w14:checkbox``).
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import TYPE_CHECKING, Any, Literal

from lxml import etree

from docx_plus.core import DocxPlusError
from docx_plus.core.ns import qn
from docx_plus.core.oxml import el, remove, sub, xpath

if TYPE_CHECKING:
    from docx.document import Document


ControlType = Literal["text", "dropdown", "combobox", "date", "checkbox"]
ControlValueT = str | bool | datetime

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
    """Raised by :func:`read_controls` when two controls share a key."""


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
        tag: The control's ``w:tag`` value (machine identifier).
        alias: The control's ``w:alias`` value (UI label), or ``None``.
        control_type: One of ``text``, ``dropdown``, ``combobox``, ``date``,
            ``checkbox``.
        value: The current value:

            - text/dropdown/combobox: ``str`` if filled, ``None`` if showing
              placeholder.
            - date: :class:`~datetime.datetime` if filled, ``None`` otherwise.
            - checkbox: always ``bool`` (no placeholder concept).

        is_placeholder: True if the control is showing its placeholder text
            (``w:showingPlcHdr`` present in sdtPr). Always ``False`` for
            checkboxes.
    """

    tag: str
    alias: str | None
    control_type: ControlType
    value: ControlValueT | None
    is_placeholder: bool


# --------------------------------------------------------------------------
# Public API.
# --------------------------------------------------------------------------


def read_controls(
    doc: Document,
    *,
    by: Literal["tag", "alias"] = "tag",
) -> dict[str, ControlValue]:
    """Return every content control in ``doc`` keyed by tag (or alias).

    Args:
        doc: The python-docx Document to inspect.
        by: Either ``"tag"`` (default) — key on ``w:tag``, every control
            included — or ``"alias"`` — key on ``w:alias``, controls without
            an alias are skipped.

    Returns:
        Mapping from key to :class:`ControlValue`.

    Raises:
        DuplicateTagError: If two controls share the same key.
    """
    out: dict[str, ControlValue] = {}
    for sdt in _iter_sdts(doc):
        info = _read_sdt(sdt)
        if info is None:
            continue
        key = info.tag if by == "tag" else info.alias
        if key is None:
            continue
        if key in out:
            raise DuplicateTagError(
                f"duplicate {by} {key!r} encountered while reading controls",
            )
        out[key] = info
    return out


def set_control_value(
    doc: Document,
    tag: str,
    value: ControlValueT,
) -> None:
    """Set the value of a control identified by ``tag``.

    Args:
        doc: The python-docx Document to modify.
        tag: The control's ``w:tag`` value.
        value: The new value. Type must match the control type:

            - text: ``str``
            - dropdown / combobox: ``str``
            - date: :class:`~datetime.datetime`
            - checkbox: ``bool``

    Raises:
        ControlNotFoundError: If no control with that tag exists.
        ControlTypeError: If ``value``'s type does not match the control type.
        ValueNotInListError: For a dropdown when ``value`` matches neither
            ``w:value`` nor ``w:displayText`` of any list item.
    """
    sdt = _find_sdt_by_tag(doc, tag)
    sdt_pr = _sdt_pr(sdt)
    sdt_content = _sdt_content(sdt)
    control_type = _classify_sdt(sdt)
    if control_type is None:
        raise ControlNotFoundError(
            f"control with tag {tag!r} has no recognised type marker",
        )

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


def clear_control(doc: Document, tag: str) -> None:
    """Reset a control to its placeholder state.

    For text/dropdown/combobox/date: re-adds ``w:showingPlcHdr`` to sdtPr and
    re-applies the ``PlaceholderText`` rStyle to every run in sdtContent. The
    placeholder text itself is preserved in place (whatever sdtContent
    currently holds).

    For checkbox: resets the checked flag to ``0`` and the glyph to
    ``☐``. Checkboxes have no placeholder mode.
    """
    sdt = _find_sdt_by_tag(doc, tag)
    sdt_pr = _sdt_pr(sdt)
    sdt_content = _sdt_content(sdt)
    control_type = _classify_sdt(sdt)
    if control_type is None:
        raise ControlNotFoundError(
            f"control with tag {tag!r} has no recognised type marker",
        )

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
    """Return the control type for an SDT, or ``None`` for unknown / rich text."""
    sdt_pr = sdt.find(qn("w:sdtPr"))
    if sdt_pr is None:
        return None
    if sdt_pr.find(qn("w:text")) is not None:
        return "text"
    if sdt_pr.find(qn("w:dropDownList")) is not None:
        return "dropdown"
    if sdt_pr.find(qn("w:comboBox")) is not None:
        return "combobox"
    if sdt_pr.find(qn("w:date")) is not None:
        return "date"
    if sdt_pr.find(qn("w14:checkbox")) is not None:
        return "checkbox"
    return None


# --------------------------------------------------------------------------
# Internals.
# --------------------------------------------------------------------------


def _iter_sdts(doc: Document) -> list[etree._Element]:
    body: Any = doc.element.body
    return [s for s in xpath(body, ".//w:sdt") if isinstance(s, etree._Element)]


def _find_sdt_by_tag(doc: Document, tag: str) -> etree._Element:
    for sdt in _iter_sdts(doc):
        sdt_pr = sdt.find(qn("w:sdtPr"))
        if sdt_pr is None:
            continue
        tag_el = sdt_pr.find(qn("w:tag"))
        if tag_el is None:
            continue
        if tag_el.get(qn("w:val")) == tag:
            return sdt
    raise ControlNotFoundError(f"no content control with tag {tag!r}")


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


def _read_sdt(sdt: etree._Element) -> ControlValue | None:
    sdt_pr = sdt.find(qn("w:sdtPr"))
    if sdt_pr is None:
        return None
    control_type = _classify_sdt(sdt)
    if control_type is None:
        return None

    tag_el = sdt_pr.find(qn("w:tag"))
    if tag_el is None:
        return None
    tag = tag_el.get(qn("w:val")) or ""

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
    tag: str,
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
    "ControlValue",
    "DuplicateTagError",
    "ValueNotInListError",
    "clear_control",
    "read_controls",
    "set_control_value",
]
