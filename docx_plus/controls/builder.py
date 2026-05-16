"""Build Word content controls (SDTs) — text, dropdown, date, checkbox.

python-docx stops at the paragraph/run layer; content controls are ``w:sdt``
elements that have to be synthesised at the lxml level. :class:`FormBuilder`
wraps a python-docx :class:`~docx.document.Document` and provides ``add_*``
methods that emit valid ``w:sdt`` blocks and append them inline to a
paragraph.

The builder handles the three failure modes the docx-forms skill prototype
identified:

1. ``w:id`` collisions — every id flows through :class:`IdRegistry`.
2. The latent ``PlaceholderText`` style — materialised on construction so the
   grey placeholder text actually renders.
3. ``w14`` namespace declaration on the document root — required by
   ``w14:checkbox``; verified at construction time.

This module imports only from ``docx_plus.core`` (SPEC §9.1). The
``PlaceholderText`` style definition is duplicated here intentionally rather
than reused from :mod:`docx_plus.styles.modify`.
"""

from __future__ import annotations

import os
from typing import TYPE_CHECKING, Any

from docx import Document
from lxml import etree

from docx_plus.core import DocxPlusError
from docx_plus.core.ids import IdRegistry
from docx_plus.core.ns import W14
from docx_plus.core.oxml import el, sub, xpath

if TYPE_CHECKING:
    from docx.document import Document as DocxDocument
    from docx.text.paragraph import Paragraph

DropdownItem = str | tuple[str, str]


# --------------------------------------------------------------------------
# Errors.
# --------------------------------------------------------------------------


class MissingNamespaceError(DocxPlusError):
    """Raised when a required namespace is not declared on the document root."""


# --------------------------------------------------------------------------
# Module constants — match Word's defaults so the rendered controls look
# right before the user touches them in Word.
# --------------------------------------------------------------------------

_PLACEHOLDER_STYLE_ID = "PlaceholderText"
_PLACEHOLDER_STYLE_NAME = "Placeholder Text"

_CHECKBOX_CHECKED_GLYPH = "☒"  # ☒
_CHECKBOX_UNCHECKED_GLYPH = "☐"  # ☐
_CHECKBOX_CHECKED_HEX = "2612"
_CHECKBOX_UNCHECKED_HEX = "2610"
_CHECKBOX_FONT = "MS Gothic"


# --------------------------------------------------------------------------
# FormBuilder.
# --------------------------------------------------------------------------


class FormBuilder:
    """Wrap a python-docx Document and add fillable content controls.

    ``self.doc`` is the underlying :class:`~docx.document.Document` — use it
    for ordinary document construction (headings, paragraphs, tables). Use the
    ``add_*`` methods to drop content controls into paragraphs you have made.

    Each ``add_*`` method appends the SDT *inline* at the end of the paragraph
    you pass, so put the field's label in the paragraph text first.
    """

    doc: DocxDocument

    def __init__(
        self,
        document_or_path: DocxDocument | str | os.PathLike[str] | None = None,
        *,
        id_registry: IdRegistry | None = None,
    ) -> None:
        """Open or wrap a document and prepare the builder state.

        Args:
            document_or_path: An open :class:`~docx.document.Document`, a path
                to a ``.docx`` file to open, or ``None`` to start a blank
                document.
            id_registry: An existing :class:`IdRegistry` to share with other
                builders. ``None`` (default) creates a fresh registry seeded
                from the document's existing SDT ids.

        Raises:
            MissingNamespaceError: If the document root does not declare the
                ``w14`` namespace (required by ``w14:checkbox``). Fresh
                python-docx documents always declare it.
        """
        if document_or_path is None:
            self.doc = Document()
        elif isinstance(document_or_path, (str, os.PathLike)):
            self.doc = Document(os.fspath(document_or_path))
        else:
            self.doc = document_or_path

        self._id_registry = id_registry if id_registry is not None else IdRegistry(self.doc)
        _verify_w14_declared(self.doc)
        _ensure_placeholder_style(self.doc)

    # -- public control builders ----------------------------------------------

    def add_text_control(
        self,
        paragraph: Paragraph,
        *,
        tag: str,
        alias: str | None = None,
        placeholder: str = "Click to enter text",
        multiline: bool = False,
    ) -> etree._Element:
        """Append an inline plain-text content control to ``paragraph``.

        Args:
            paragraph: The python-docx paragraph to append into.
            tag: Stable machine-readable identifier for the control.
            alias: Optional human-friendly label shown in Word's UI.
            placeholder: The grey "click here" prompt rendered inside the
                empty control.
            multiline: If ``True``, allow hard line breaks inside the control
                (use for addresses, comment boxes).

        Returns:
            The created ``w:sdt`` element.
        """
        sdt, sdt_pr, sdt_content = self._new_sdt(tag=tag, alias=alias)
        sub(sdt_pr, "w:showingPlcHdr")
        text_attrs: dict[str, str] = {"w:multiLine": "1"} if multiline else {}
        sub(sdt_pr, "w:text", **text_attrs)

        sdt_content.append(_placeholder_run(placeholder))
        sdt.append(sdt_content)
        paragraph._p.append(sdt)
        return sdt

    def add_dropdown(
        self,
        paragraph: Paragraph,
        *,
        tag: str,
        items: list[DropdownItem],
        alias: str | None = None,
        placeholder: str = "Choose an item",
        editable: bool = False,
    ) -> etree._Element:
        """Append a dropdown (or combobox) content control to ``paragraph``.

        Args:
            paragraph: The python-docx paragraph to append into.
            tag: Stable machine-readable identifier for the control.
            items: A list of either plain strings, or ``(display, value)``
                tuples when the stored value should differ from the shown
                label.
            alias: Optional human-friendly label shown in Word's UI.
            placeholder: The "Choose an item" prompt rendered inside the
                empty control. A placeholder list-item with empty value is
                also added as the first dropdown entry.
            editable: If ``True``, produce a ``w:comboBox`` (user may type a
                value not in the list) instead of a ``w:dropDownList``.

        Returns:
            The created ``w:sdt`` element.

        Raises:
            TypeError: If ``items`` contains anything that is not a string
                or a 2-tuple of strings.
        """
        sdt, sdt_pr, sdt_content = self._new_sdt(tag=tag, alias=alias)
        sub(sdt_pr, "w:showingPlcHdr")
        list_tag = "w:comboBox" if editable else "w:dropDownList"
        list_el = sub(sdt_pr, list_tag)

        sub(list_el, "w:listItem", **{"w:displayText": placeholder, "w:value": ""})
        for raw_item in items:
            display, value = _normalise_dropdown_item(raw_item)
            sub(list_el, "w:listItem", **{"w:displayText": display, "w:value": value})

        sdt_content.append(_placeholder_run(placeholder))
        sdt.append(sdt_content)
        paragraph._p.append(sdt)
        return sdt

    def add_date_picker(
        self,
        paragraph: Paragraph,
        *,
        tag: str,
        alias: str | None = None,
        placeholder: str = "Click to select a date",
        date_format: str = "M/d/yyyy",
        lcid: str = "en-US",
    ) -> etree._Element:
        """Append a date-picker content control to ``paragraph``.

        Args:
            paragraph: The python-docx paragraph to append into.
            tag: Stable machine-readable identifier for the control.
            alias: Optional human-friendly label shown in Word's UI.
            placeholder: The grey "click here" prompt rendered inside the
                empty control.
            date_format: Word's date-format string (e.g. ``"M/d/yyyy"``,
                ``"dddd, MMMM d, yyyy"``).
            lcid: Locale identifier (BCP-47 form, e.g. ``"en-US"``).

        Returns:
            The created ``w:sdt`` element.
        """
        sdt, sdt_pr, sdt_content = self._new_sdt(tag=tag, alias=alias)
        sub(sdt_pr, "w:showingPlcHdr")
        date_el = sub(sdt_pr, "w:date")
        sub(date_el, "w:dateFormat", **{"w:val": date_format})
        sub(date_el, "w:lid", **{"w:val": lcid})
        sub(date_el, "w:storeMappedDataAs", **{"w:val": "dateTime"})
        sub(date_el, "w:calendar", **{"w:val": "gregorian"})

        sdt_content.append(_placeholder_run(placeholder))
        sdt.append(sdt_content)
        paragraph._p.append(sdt)
        return sdt

    def add_checkbox(
        self,
        paragraph: Paragraph,
        *,
        tag: str,
        alias: str | None = None,
        checked: bool = False,
    ) -> etree._Element:
        """Append a Word 2010+ ``w14:checkbox`` content control to ``paragraph``.

        The visible glyph and the ``w14:checked`` flag are kept in sync, so
        the box renders correctly even before Word ever opens the file.

        Args:
            paragraph: The python-docx paragraph to append into.
            tag: Stable machine-readable identifier for the control.
            alias: Optional human-friendly label shown in Word's UI.
            checked: Initial checked state.

        Returns:
            The created ``w:sdt`` element.
        """
        sdt, sdt_pr, sdt_content = self._new_sdt(tag=tag, alias=alias)
        checkbox = sub(sdt_pr, "w14:checkbox")
        sub(checkbox, "w14:checked", **{"w14:val": "1" if checked else "0"})
        sub(
            checkbox,
            "w14:checkedState",
            **{"w14:val": _CHECKBOX_CHECKED_HEX, "w14:font": _CHECKBOX_FONT},
        )
        sub(
            checkbox,
            "w14:uncheckedState",
            **{"w14:val": _CHECKBOX_UNCHECKED_HEX, "w14:font": _CHECKBOX_FONT},
        )

        sdt_content.append(
            _checkbox_glyph_run(
                _CHECKBOX_CHECKED_GLYPH if checked else _CHECKBOX_UNCHECKED_GLYPH,
            ),
        )
        sdt.append(sdt_content)
        paragraph._p.append(sdt)
        return sdt

    def save(self, path: str | os.PathLike[str]) -> str:
        """Save the wrapped document to ``path`` and return the path as a string."""
        self.doc.save(os.fspath(path))
        return os.fspath(path)

    # -- internals ------------------------------------------------------------

    def _new_sdt(
        self,
        *,
        tag: str,
        alias: str | None,
    ) -> tuple[etree._Element, etree._Element, etree._Element]:
        """Build the shared ``w:sdt``/``w:sdtPr``/``w:sdtContent`` scaffold.

        sdtPr child order matches the docx-forms skill prototype:
        ``[alias?], tag, id, [showingPlcHdr], <type-marker>``. Caller appends
        showingPlcHdr and the type marker (and finally the populated
        sdtContent) in that order.
        """
        sdt = el("w:sdt")
        sdt_pr = sub(sdt, "w:sdtPr")

        if alias is not None:
            sub(sdt_pr, "w:alias", **{"w:val": alias})
        sub(sdt_pr, "w:tag", **{"w:val": tag})
        sub(sdt_pr, "w:id", **{"w:val": str(self._id_registry.next())})

        sdt_content = el("w:sdtContent")
        return sdt, sdt_pr, sdt_content


# --------------------------------------------------------------------------
# Module-level helpers (private).
# --------------------------------------------------------------------------


def _placeholder_run(text: str) -> etree._Element:
    """Build a ``w:r`` carrying the ``PlaceholderText`` rStyle and ``text``."""
    run = el("w:r")
    rpr = sub(run, "w:rPr")
    sub(rpr, "w:rStyle", **{"w:val": _PLACEHOLDER_STYLE_ID})
    text_el = sub(run, "w:t")
    text_el.text = text
    return run


def _checkbox_glyph_run(glyph: str) -> etree._Element:
    """Build the ``w:r`` that renders the checkbox glyph in ``MS Gothic``."""
    run = el("w:r")
    rpr = sub(run, "w:rPr")
    sub(
        rpr,
        "w:rFonts",
        **{
            "w:ascii": _CHECKBOX_FONT,
            "w:hAnsi": _CHECKBOX_FONT,
            "w:eastAsia": _CHECKBOX_FONT,
        },
    )
    text_el = sub(run, "w:t")
    text_el.text = glyph
    return run


def _normalise_dropdown_item(raw: DropdownItem) -> tuple[str, str]:
    """Convert ``raw`` into a ``(display, value)`` pair."""
    if isinstance(raw, str):
        return raw, raw
    if isinstance(raw, tuple) and len(raw) == 2 and all(isinstance(p, str) for p in raw):
        return raw[0], raw[1]
    raise TypeError(
        f"dropdown item must be str or (str, str) tuple; got {type(raw).__name__}: {raw!r}",
    )


def _verify_w14_declared(doc: DocxDocument) -> None:
    """Raise :class:`MissingNamespaceError` if the document root lacks ``w14``."""
    nsmap: dict[str | None, str] = doc.element.nsmap
    declared = any(uri == W14 for uri in nsmap.values())
    if not declared:
        raise MissingNamespaceError(
            "document root does not declare the w14 namespace; "
            "w14:checkbox controls cannot be authored. Expected nsmap entry "
            f"with URI {W14!r}.",
        )


def _ensure_placeholder_style(doc: DocxDocument) -> None:
    """Materialise the ``PlaceholderText`` character style if absent.

    Independent of :func:`docx_plus.styles.modify.ensure_style` per SPEC §9.1
    (controls/ may not import styles/). The definition mirrors Word's default.
    """
    styles_root: Any = doc.styles.element
    for style in xpath(styles_root, f"./w:style[@w:styleId='{_PLACEHOLDER_STYLE_ID}']"):
        if isinstance(style, etree._Element):
            return

    style_el = el(
        "w:style",
        **{"w:type": "character", "w:styleId": _PLACEHOLDER_STYLE_ID},
    )
    sub(style_el, "w:name", **{"w:val": _PLACEHOLDER_STYLE_NAME})
    sub(style_el, "w:basedOn", **{"w:val": "DefaultParagraphFont"})
    sub(style_el, "w:uiPriority", **{"w:val": "99"})
    sub(style_el, "w:semiHidden")
    sub(style_el, "w:unhideWhenUsed")
    rpr = sub(style_el, "w:rPr")
    sub(rpr, "w:color", **{"w:val": "808080"})
    styles_root.append(style_el)


__all__ = ["DropdownItem", "FormBuilder", "MissingNamespaceError"]
