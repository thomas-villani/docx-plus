"""Style creation, modification, application, deletion — Word-native workflow.

The companion to :mod:`docx_plus.styles.inspect`. Where ``inspect`` reads the
cascade, this module writes the styles that drive it. The intent is to push
formatting changes through *style definitions* rather than scattering direct
formatting; SPEC §5 calls this the "Word-native" workflow.

Property kwargs accepted by :func:`create_style` and :func:`modify_style` use
the same field names as :class:`~docx_plus.styles.inspect.ResolvedFormatting`
so a value resolved by the inspector can be round-tripped back through the
modifier without translation. Schema-strict child ordering for ``CT_Style``,
``CT_PPr``, and ``CT_RPr`` is enforced internally so the produced styles.xml
matches what Word writes.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, Literal

from lxml import etree

from docx_plus.core import DocxPlusError
from docx_plus.core.ns import qn
from docx_plus.core.oxml import el, xpath

if TYPE_CHECKING:
    from docx.document import Document
    from docx.table import _Cell
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run


StyleType = Literal["paragraph", "character", "table", "numbering"]


# --------------------------------------------------------------------------
# Errors.
# --------------------------------------------------------------------------


class StyleExistsError(DocxPlusError):
    """Raised by :func:`create_style` when the style id is already defined."""


class StyleNotFoundError(DocxPlusError):
    """Raised when an operation references a style id that does not exist."""


class StyleInUseError(DocxPlusError):
    """Raised by :func:`delete_style` when the style is referenced and ``force=False``."""


class UnknownStylePropertyError(DocxPlusError, TypeError):
    """Raised when a property kwarg is not a recognised style property."""


# --------------------------------------------------------------------------
# Schema-correct child orderings.
# --------------------------------------------------------------------------
# These are the local-name sequences for the relevant CT_* types in the
# WordprocessingML schema. Inserting children in these orders prevents the
# "Word repaired the file" failure mode (silent reorder by the renderer) and
# matches what Word itself writes. See ECMA-376 17.7 (style hierarchy) and
# 17.3.1 (paragraph properties) / 17.3.2 (run properties).

_STYLE_CHILD_ORDER: tuple[str, ...] = (
    "name",
    "aliases",
    "basedOn",
    "next",
    "link",
    "autoRedefine",
    "hidden",
    "uiPriority",
    "semiHidden",
    "unhideWhenUsed",
    "qFormat",
    "locked",
    "personal",
    "personalCompose",
    "personalReply",
    "rsid",
    "pPr",
    "rPr",
    "tblPr",
    "trPr",
    "tcPr",
    "tblStylePr",
)

_PPR_CHILD_ORDER: tuple[str, ...] = (
    "pStyle",
    "keepNext",
    "keepLines",
    "pageBreakBefore",
    "framePr",
    "widowControl",
    "numPr",
    "suppressLineNumbers",
    "pBdr",
    "shd",
    "tabs",
    "suppressAutoHyphens",
    "kinsoku",
    "wordWrap",
    "overflowPunct",
    "topLinePunct",
    "autoSpaceDE",
    "autoSpaceDN",
    "bidi",
    "adjustRightInd",
    "snapToGrid",
    "spacing",
    "ind",
    "contextualSpacing",
    "mirrorIndents",
    "suppressOverlap",
    "jc",
    "textDirection",
    "textAlignment",
    "textboxTightWrap",
    "outlineLvl",
    "divId",
    "cnfStyle",
    "rPr",
    "sectPr",
    "pPrChange",
)

_RPR_CHILD_ORDER: tuple[str, ...] = (
    "rStyle",
    "rFonts",
    "b",
    "bCs",
    "i",
    "iCs",
    "caps",
    "smallCaps",
    "strike",
    "dstrike",
    "outline",
    "shadow",
    "emboss",
    "imprint",
    "noProof",
    "snapToGrid",
    "vanish",
    "webHidden",
    "color",
    "spacing",
    "w",
    "kern",
    "position",
    "sz",
    "szCs",
    "highlight",
    "u",
    "effect",
    "bdr",
    "shd",
    "fitText",
    "vertAlign",
    "rtl",
    "cs",
    "em",
    "lang",
    "eastAsianLayout",
    "specVanish",
    "oMath",
    "rPrChange",
)


# --------------------------------------------------------------------------
# StyleInfo / StyleProxy — read-side surfaces.
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class StyleInfo:
    """Lightweight summary of a style for :func:`list_styles`.

    Attributes:
        style_id: The ``w:styleId`` value (machine identifier).
        name: The ``w:name`` value (human-readable display name).
        style_type: ``"paragraph"`` | ``"character"`` | ``"table"`` |
            ``"numbering"``.
        based_on: The id of the parent style in the basedOn chain, if any.
        is_default: True if the style carries ``w:default="1"``.
        is_latent: True only for entries returned with
            ``include_latent=True`` that aren't materialised in styles.xml.
    """

    style_id: str
    name: str
    style_type: StyleType
    based_on: str | None = None
    is_default: bool = False
    is_latent: bool = False


class StyleProxy:
    """Lightweight live wrapper around a ``<w:style>`` element.

    The proxy holds a reference to the live element rather than a snapshot —
    reads always reflect current state. Mutating methods delegate to
    :func:`modify_style` so child-element ordering and toggle semantics stay
    consistent. The :attr:`element` attribute is an escape hatch for callers
    that need direct lxml access (SPEC §5).
    """

    def __init__(self, doc: Document, element: etree._Element) -> None:
        """Wrap an existing ``<w:style>`` element in ``doc``."""
        self._doc = doc
        self.element = element

    @property
    def style_id(self) -> str:
        """The style's ``w:styleId`` (machine identifier)."""
        sid = self.element.get(qn("w:styleId"))
        if sid is None:
            raise StyleNotFoundError("style element is missing w:styleId")
        return sid

    @property
    def style_type(self) -> StyleType:
        """The style's ``w:type`` (paragraph/character/table/numbering)."""
        raw = self.element.get(qn("w:type")) or "paragraph"
        if raw not in ("paragraph", "character", "table", "numbering"):
            raise StyleNotFoundError(f"unknown w:type {raw!r}")
        # ``raw`` is narrowed to the StyleType literal by the check above, but
        # mypy doesn't propagate the narrowing through a tuple membership test.
        return raw  # type: ignore[return-value]

    @property
    def name(self) -> str | None:
        """The style's display name from ``w:name``."""
        name_el = self.element.find(qn("w:name"))
        return name_el.get(qn("w:val")) if name_el is not None else None

    @property
    def based_on(self) -> str | None:
        """The style id this one inherits from via ``w:basedOn``, if any."""
        based = self.element.find(qn("w:basedOn"))
        return based.get(qn("w:val")) if based is not None else None

    @property
    def next_style(self) -> str | None:
        """The style id Word applies to the paragraph after one styled with this."""
        nxt = self.element.find(qn("w:next"))
        return nxt.get(qn("w:val")) if nxt is not None else None

    @property
    def linked_style(self) -> str | None:
        """The companion character style id from ``w:link``, if any."""
        lnk = self.element.find(qn("w:link"))
        return lnk.get(qn("w:val")) if lnk is not None else None

    @property
    def ui_priority(self) -> int | None:
        """The Word style-gallery sort priority from ``w:uiPriority``."""
        pri = self.element.find(qn("w:uiPriority"))
        if pri is None:
            return None
        raw = pri.get(qn("w:val"))
        if raw is None:
            return None
        try:
            return int(raw)
        except ValueError:
            return None

    @property
    def q_format(self) -> bool:
        """True if the style is shown in Word's quick-style gallery."""
        return self.element.find(qn("w:qFormat")) is not None

    def modify(self, **properties: Any) -> StyleProxy:
        """Convenience: thin wrapper around :func:`modify_style`."""
        return modify_style(self._doc, self.style_id, **properties)

    def delete(self, *, force: bool = False) -> None:
        """Convenience: thin wrapper around :func:`delete_style`."""
        delete_style(self._doc, self.style_id, force=force)

    def __repr__(self) -> str:
        """Compact repr for diagnostics."""
        return f"StyleProxy(style_id={self.style_id!r}, type={self.style_type!r})"


# --------------------------------------------------------------------------
# Property writer registry.
# --------------------------------------------------------------------------
# Each property maps to a write strategy. A "writer" is a callable that, given
# a style element and the user value, mutates the style's pPr/rPr appropriately
# while honouring schema order and merge semantics.

_PARAGRAPH_LEVEL_PROPS: frozenset[str] = frozenset(
    {
        "alignment",
        "indent_left",
        "indent_right",
        "indent_first_line",
        "spacing_before",
        "spacing_after",
        "line_spacing",
        "line_spacing_rule",
        "keep_with_next",
        "keep_lines",
        "page_break_before",
        "outline_level",
    }
)

_RUN_LEVEL_PROPS: frozenset[str] = frozenset(
    {
        "font_name",
        "font_size",
        "bold",
        "italic",
        "underline",
        "strike",
        "color_rgb",
        "highlight",
        "caps",
        "small_caps",
        "vanish",
        "vert_align",
    }
)

_TOGGLE_PROPS: dict[str, str] = {
    "bold": "b",
    "italic": "i",
    "caps": "caps",
    "small_caps": "smallCaps",
    "strike": "strike",
    "vanish": "vanish",
}

_ALL_PROPS: frozenset[str] = _PARAGRAPH_LEVEL_PROPS | _RUN_LEVEL_PROPS


# --------------------------------------------------------------------------
# Public API: create / modify / apply / delete / ensure / list.
# --------------------------------------------------------------------------


def create_style(
    doc: Document,
    style_id: str,
    *,
    style_type: StyleType = "paragraph",
    name: str | None = None,
    based_on: str | None = None,
    next_style: str | None = None,
    linked_style: str | None = None,
    ui_priority: int = 99,
    q_format: bool = False,
    custom: bool = True,
    **properties: Any,
) -> StyleProxy:
    """Define a new style in ``doc``.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to mutate.
        style_id: Machine identifier (``w:styleId``). Must be unique within
            the document's styles.xml.
        style_type: ``"paragraph"`` (default), ``"character"``, ``"table"``,
            or ``"numbering"``.
        name: Display name. Defaults to ``style_id``.
        based_on: Parent style id in the basedOn chain.
        next_style: Style applied to the paragraph that follows one styled
            with this style (e.g. ``Heading1`` -> ``Normal``).
        linked_style: Companion character style for a paragraph style (Word's
            ``Heading1`` <-> ``Heading1Char`` pairing).
        ui_priority: Sort priority in Word's style gallery (lower = higher).
        q_format: Show in Word's quick-style gallery if True.
        custom: Mark as a custom style (``w:customStyle="1"``). Defaults True
            because user-defined styles should not be confused with built-ins.
        **properties: Any field name from :class:`ResolvedFormatting`. See
            module docstring for the supported set.

    Returns:
        A :class:`StyleProxy` for the new style.

    Raises:
        StyleExistsError: If ``style_id`` is already defined.
        UnknownStylePropertyError: If ``**properties`` contains a key that is
            not a recognised style property.
    """
    _validate_property_keys(properties)
    styles_root = doc.styles.element
    if _find_style_element(styles_root, style_id) is not None:
        raise StyleExistsError(f"style {style_id!r} already exists")

    style_el = el("w:style", **{"w:type": style_type, "w:styleId": style_id})
    if custom:
        style_el.set(qn("w:customStyle"), "1")

    _set_simple_child(style_el, "name", {"w:val": name or style_id})
    if based_on is not None:
        _set_simple_child(style_el, "basedOn", {"w:val": based_on})
    if next_style is not None:
        _set_simple_child(style_el, "next", {"w:val": next_style})
    if linked_style is not None:
        _set_simple_child(style_el, "link", {"w:val": linked_style})
    _set_simple_child(style_el, "uiPriority", {"w:val": str(ui_priority)})
    if q_format:
        _set_simple_child(style_el, "qFormat", {})

    for prop_name, value in properties.items():
        _write_property(style_el, prop_name, value)

    styles_root.append(style_el)
    return StyleProxy(doc, style_el)


def modify_style(
    doc: Document,
    style_id: str,
    *,
    if_missing: Literal["raise", "create"] = "raise",
    **properties: Any,
) -> StyleProxy:
    """Update an existing style's properties in place.

    Pass only the properties to change; others are preserved. Per SPEC §5,
    toggle properties (``bold``, ``italic``, …) treat ``True``/``False`` as
    explicit settings (writing ``w:val="true"``/``"false"``) and ``None`` as
    "clear the setting so XOR with the parent resumes". Non-toggle properties
    treat ``None`` as "remove this property".

    Args:
        doc: The document containing the style.
        style_id: Identifier of the style to modify.
        if_missing: ``"raise"`` (default) raises :class:`StyleNotFoundError`
            when the style is not defined; ``"create"`` falls through to
            :func:`create_style` with the supplied properties as the initial
            definition.
        **properties: Any field name from :class:`ResolvedFormatting`.

    Returns:
        A :class:`StyleProxy` for the modified style.

    Raises:
        StyleNotFoundError: If ``style_id`` is undefined and
            ``if_missing="raise"``.
        UnknownStylePropertyError: If ``**properties`` contains an
            unrecognised key.
    """
    _validate_property_keys(properties)
    styles_root = doc.styles.element
    style_el = _find_style_element(styles_root, style_id)
    if style_el is None:
        if if_missing == "create":
            return create_style(doc, style_id, **properties)
        raise StyleNotFoundError(f"style {style_id!r} is not defined")
    for prop_name, value in properties.items():
        _write_property(style_el, prop_name, value)
    return StyleProxy(doc, style_el)


def apply_style(target: Paragraph | Run | _Cell, style_id: str) -> None:
    """Apply a style by id to a paragraph, run, or cell.

    Resolves the style id against the target's owning document and writes the
    appropriate ``w:pStyle`` (paragraph), ``w:rStyle`` (run), or — for cells —
    sets ``w:pStyle`` on every paragraph in the cell.

    Args:
        target: A python-docx :class:`~docx.text.paragraph.Paragraph`,
            :class:`~docx.text.run.Run`, or :class:`~docx.table._Cell`.
        style_id: The id of an already-defined style.

    Raises:
        StyleNotFoundError: If ``style_id`` is not defined in the target's
            document.
    """
    from docx.table import _Cell as _CellCls
    from docx.text.paragraph import Paragraph as _Paragraph
    from docx.text.run import Run as _Run

    if not isinstance(target, (_Paragraph, _Run, _CellCls)):
        kind = type(target).__name__
        raise TypeError(f"apply_style expects Paragraph, Run, or _Cell; got {kind}")

    part: Any = target.part
    doc = part.document
    if _find_style_element(doc.styles.element, style_id) is None:
        raise StyleNotFoundError(f"style {style_id!r} is not defined")

    if isinstance(target, _Paragraph):
        _set_paragraph_style(target._p, style_id)
    elif isinstance(target, _Run):
        _set_run_style(target._r, style_id)
    else:
        for p in target.paragraphs:
            _set_paragraph_style(p._p, style_id)


def ensure_style(
    doc: Document,
    style_id: str,
    *,
    match_existing: bool = False,
    **defaults_if_creating: Any,
) -> StyleProxy:
    """Idempotent style materialisation.

    If ``style_id`` is already defined, return a proxy without modifying it.
    If it names a known built-in (Heading1, Title, ListParagraph, …),
    materialise it from the built-in table — Word's defaults, not
    ``defaults_if_creating``. Otherwise create a custom style with
    ``defaults_if_creating`` as the initial properties.

    Args:
        doc: Document to ensure the style on.
        style_id: Identifier to ensure.
        match_existing: If True, before falling back to creation, search for
            an existing style whose ``w:styleId`` or ``w:name`` matches
            ``style_id`` case/space-insensitively (via
            :func:`find_matching_style`). If found, that proxy is returned —
            note its ``style_id`` may differ from the requested one, so
            callers using ``apply_style`` should pass ``proxy.style_id``.
            For document-wide normalisation, use :func:`remap_styles` which
            rewrites body references too.
        **defaults_if_creating: Properties to use when creating a *custom*
            style. Ignored when materialising a known built-in (the user
            asked for the built-in; Word's defaults are what matters).

    Returns:
        A :class:`StyleProxy` for the existing-or-newly-created style.
    """
    styles_root = doc.styles.element
    existing = _find_style_element(styles_root, style_id)
    if existing is not None:
        return StyleProxy(doc, existing)
    if match_existing:
        matched_id = find_matching_style(doc, style_id)
        if matched_id is not None:
            matched_el = _find_style_element(styles_root, matched_id)
            if matched_el is not None:
                return StyleProxy(doc, matched_el)
    builtin = _BUILTIN_STYLES.get(style_id)
    if builtin is not None:
        return _materialise_builtin(doc, style_id, builtin)
    return create_style(doc, style_id, **defaults_if_creating)


def find_matching_style(doc: Document, target_id: str) -> str | None:
    """Find an existing style that fulfils the role of ``target_id``.

    Matches case/space-insensitively against both the ``w:styleId`` and
    ``w:name`` of every defined style. Useful when a document uses a renamed
    or differently-cased version of a built-in style (``"Heading 1"`` with a
    space, ``"heading1"`` lower-case, …).

    Args:
        doc: Document to search.
        target_id: The id you want to map onto (e.g. ``"Heading1"``).

    Returns:
        The :attr:`w:styleId` of the first matching defined style, or
        ``None`` if none match. If a style with id ``target_id`` is already
        defined exactly, returns ``target_id`` (the trivial match).
    """
    target_norm = _normalize_style_key(target_id)
    if not target_norm:
        return None
    styles_root = doc.styles.element
    for style_el in styles_root.findall(qn("w:style")):
        sid: str | None = style_el.get(qn("w:styleId"))
        if sid is None:
            continue
        if _normalize_style_key(sid) == target_norm:
            return sid
        name_el = style_el.find(qn("w:name"))
        if name_el is not None:
            name: str | None = name_el.get(qn("w:val"))
            if name is not None and _normalize_style_key(name) == target_norm:
                return sid
    return None


def remap_styles(
    doc: Document,
    *,
    targets: list[str] | None = None,
    mapping: dict[str, str] | None = None,
    create_missing: bool = False,
) -> dict[str, str]:
    """Reconcile a doc's styles against a set of canonical ids.

    For each id in ``targets``, resolve it to an existing style by:

    1. Exact match — the id is already defined in ``styles.xml``.
    2. The supplied ``mapping`` if it names this target.
    3. :func:`find_matching_style` (case/space-insensitive on id and name).
    4. If ``create_missing=True`` and the target is in the known built-ins
       table, materialise it from the table — which declares
       ``basedOn="Normal"`` so the new style inherits the doc's customised
       Normal (fonts, colours, …) automatically.

    Body references (``w:pStyle``, ``w:rStyle``, ``w:tblStyle``) pointing at
    the original target id are rewritten in place to the resolved id, so a
    subsequent :func:`apply_style` works without further translation. Refs
    *between* styles in ``styles.xml`` (``basedOn``, ``next``, ``link``) are
    left untouched — this keeps the remap a non-destructive rewrite.

    Args:
        doc: Document to remap.
        targets: Ids to reconcile. Defaults to every entry in the
            known-built-ins table.
        mapping: Optional explicit ``{target: existing_id}`` overrides
            applied before the matcher.
        create_missing: If True, fall back to materialising from the
            built-ins table when nothing else matches. Only works for ids
            that have a built-in entry.

    Returns:
        ``{target_id: resolved_id}`` for every target resolved. When
        ``target_id == resolved_id``, the doc already had it or it was just
        created. Targets unresolved after all four steps are omitted.

    Raises:
        StyleNotFoundError: If ``mapping`` names an ``existing_id`` that is
            not defined in the document.
    """
    styles_root = doc.styles.element
    target_ids: list[str] = list(targets) if targets is not None else list(_BUILTIN_STYLES)
    explicit: dict[str, str] = dict(mapping) if mapping is not None else {}

    for target, existing_id in explicit.items():
        if _find_style_element(styles_root, existing_id) is None:
            raise StyleNotFoundError(
                f"mapping for {target!r} points at undefined style {existing_id!r}"
            )

    resolved: dict[str, str] = {}
    for target_id in target_ids:
        if _find_style_element(styles_root, target_id) is not None:
            resolved[target_id] = target_id
            continue
        if target_id in explicit:
            resolved[target_id] = explicit[target_id]
            continue
        match = find_matching_style(doc, target_id)
        if match is not None:
            resolved[target_id] = match
            continue
        if create_missing and target_id in _BUILTIN_STYLES:
            _materialise_builtin(doc, target_id, _BUILTIN_STYLES[target_id])
            resolved[target_id] = target_id
            continue
        # Unresolved — omit from result.

    body_root: Any = doc.part.element
    for target_id, resolved_id in resolved.items():
        if target_id == resolved_id:
            continue
        for tag in ("pStyle", "rStyle", "tblStyle"):
            for ref in xpath(body_root, f"//w:{tag}[@w:val='{target_id}']"):
                if isinstance(ref, etree._Element):
                    ref.set(qn("w:val"), resolved_id)
    return resolved


def _normalize_style_key(value: str) -> str:
    """Lowercase + whitespace-stripped key for case/space-insensitive match."""
    return "".join(ch for ch in value.lower() if not ch.isspace())


def list_styles(
    doc: Document,
    *,
    style_type: StyleType | None = None,
    include_latent: bool = False,
) -> list[StyleInfo]:
    """List defined styles in ``doc``.

    Args:
        doc: Document to inspect.
        style_type: If given, restrict to styles of this ``w:type``.
        include_latent: If True, also yield built-ins from the known-built-ins
            table that aren't materialised in ``styles.xml``. Latent entries
            have ``is_latent=True`` so callers can distinguish them.

    Returns:
        A list of :class:`StyleInfo`, materialised styles first, latent
        entries (if requested) afterwards. Order within each group is the
        order they appear in ``styles.xml`` (or the built-ins table).
    """
    styles_root = doc.styles.element
    materialised: list[StyleInfo] = []
    seen_ids: set[str] = set()
    for style_el in styles_root.findall(qn("w:style")):
        info = _style_info_from_element(style_el)
        if info is None:
            continue
        if style_type is not None and info.style_type != style_type:
            continue
        materialised.append(info)
        seen_ids.add(info.style_id)
    if not include_latent:
        return materialised
    latent: list[StyleInfo] = []
    for sid, spec in _BUILTIN_STYLES.items():
        if sid in seen_ids:
            continue
        spec_type: StyleType = spec.get("style_type", "paragraph")
        if style_type is not None and spec_type != style_type:
            continue
        latent.append(
            StyleInfo(
                style_id=sid,
                name=spec.get("name", sid),
                style_type=spec_type,
                based_on=spec.get("based_on"),
                is_latent=True,
            )
        )
    return materialised + latent


def delete_style(doc: Document, style_id: str, *, force: bool = False) -> None:
    """Remove a style definition from ``doc``.

    Args:
        doc: Document containing the style.
        style_id: Identifier of the style to remove.
        force: If False (default), refuse to delete a style referenced by any
            paragraph, run, table, or other style. If True, delete anyway —
            Word will fall back to ``Normal`` for orphaned references.

    Raises:
        StyleNotFoundError: If ``style_id`` is not defined.
        StyleInUseError: If ``force=False`` and the style has any references.
    """
    styles_root = doc.styles.element
    style_el = _find_style_element(styles_root, style_id)
    if style_el is None:
        raise StyleNotFoundError(f"style {style_id!r} is not defined")
    if not force:
        refs = _find_references(doc, style_id)
        if refs:
            raise StyleInUseError(
                f"style {style_id!r} is referenced ({len(refs)} place(s)); "
                "pass force=True to delete anyway"
            )
    styles_root.remove(style_el)


# --------------------------------------------------------------------------
# Property writers — schema-aware mutation of pPr / rPr inside a style.
# --------------------------------------------------------------------------


def _validate_property_keys(properties: dict[str, Any]) -> None:
    unknown = set(properties) - _ALL_PROPS
    if unknown:
        names = ", ".join(sorted(unknown))
        raise UnknownStylePropertyError(f"unknown style properties: {names}")


def _write_property(style_el: etree._Element, name: str, value: Any) -> None:
    """Dispatch a (name, value) pair to the right pPr/rPr writer."""
    if name in _PARAGRAPH_LEVEL_PROPS:
        _write_paragraph_property(style_el, name, value)
    elif name in _RUN_LEVEL_PROPS:
        _write_run_property(style_el, name, value)
    else:
        raise UnknownStylePropertyError(f"unknown style property {name!r}")


def _write_paragraph_property(style_el: etree._Element, name: str, value: Any) -> None:
    if name == "alignment":
        _write_simple_pchild(style_el, "jc", value)
    elif name == "outline_level":
        _write_simple_pchild(style_el, "outlineLvl", str(value) if value is not None else None)
    elif name in ("keep_with_next", "keep_lines", "page_break_before"):
        tag = {
            "keep_with_next": "keepNext",
            "keep_lines": "keepLines",
            "page_break_before": "pageBreakBefore",
        }[name]
        _write_bool_flag(style_el, "ppr", tag, value)
    elif name in ("indent_left", "indent_right", "indent_first_line"):
        _write_indent(style_el, name, value)
    elif name in ("spacing_before", "spacing_after", "line_spacing", "line_spacing_rule"):
        _write_spacing(style_el, name, value)
    else:
        raise UnknownStylePropertyError(f"unhandled paragraph property {name!r}")


def _write_run_property(style_el: etree._Element, name: str, value: Any) -> None:
    if name in _TOGGLE_PROPS:
        _write_toggle(style_el, _TOGGLE_PROPS[name], value)
    elif name == "font_size":
        if value is None:
            _remove_rpr_child(style_el, "sz")
            _remove_rpr_child(style_el, "szCs")
        else:
            half_pts = str(int(round(float(value) * 2)))
            _set_rpr_child(style_el, "sz", {"w:val": half_pts})
            _set_rpr_child(style_el, "szCs", {"w:val": half_pts})
    elif name == "font_name":
        _write_font_name(style_el, value)
    elif name == "color_rgb":
        _write_color(style_el, value)
    elif name == "underline":
        _write_simple_rchild(style_el, "u", value)
    elif name == "highlight":
        _write_simple_rchild(style_el, "highlight", value)
    elif name == "vert_align":
        _write_simple_rchild(style_el, "vertAlign", value)
    else:
        raise UnknownStylePropertyError(f"unhandled run property {name!r}")


def _write_simple_pchild(style_el: etree._Element, tag: str, value: str | None) -> None:
    if value is None:
        _remove_ppr_child(style_el, tag)
    else:
        _set_ppr_child(style_el, tag, {"w:val": str(value)})


def _write_simple_rchild(style_el: etree._Element, tag: str, value: str | None) -> None:
    if value is None:
        _remove_rpr_child(style_el, tag)
    else:
        _set_rpr_child(style_el, tag, {"w:val": str(value)})


def _write_bool_flag(
    style_el: etree._Element, container: Literal["ppr", "rpr"], tag: str, value: bool | None
) -> None:
    """Write a presence-or-val=false boolean flag into pPr or rPr."""
    if value is None:
        _remove_pr_child(style_el, container, tag)
        return
    attrs = {"w:val": "false"} if value is False else {}
    if container == "ppr":
        _set_ppr_child(style_el, tag, attrs)
    else:
        _set_rpr_child(style_el, tag, attrs)


def _write_toggle(style_el: etree._Element, tag: str, value: bool | None) -> None:
    """Toggle property writer: True→presence, False→val=false, None→remove."""
    if value is None:
        _remove_rpr_child(style_el, tag)
        return
    attrs = {"w:val": "false"} if value is False else {}
    _set_rpr_child(style_el, tag, attrs)


def _write_indent(style_el: etree._Element, prop: str, value: int | None) -> None:
    """Merge into <w:ind/>; supports left/right/firstLine/hanging via prop name."""
    ppr = _get_or_create_ppr(style_el)
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = el("w:ind")
        _ordered_insert(ppr, ind, _PPR_CHILD_ORDER)
    if prop == "indent_left":
        _set_or_clear_attr(ind, "w:left", value)
    elif prop == "indent_right":
        _set_or_clear_attr(ind, "w:right", value)
    elif prop == "indent_first_line":
        # Negative value -> hanging; positive -> firstLine. Clear the other.
        if value is None:
            _set_or_clear_attr(ind, "w:firstLine", None)
            _set_or_clear_attr(ind, "w:hanging", None)
        elif value < 0:
            _set_or_clear_attr(ind, "w:firstLine", None)
            _set_or_clear_attr(ind, "w:hanging", -int(value))
        else:
            _set_or_clear_attr(ind, "w:hanging", None)
            _set_or_clear_attr(ind, "w:firstLine", int(value))
    if not ind.attrib:
        ppr.remove(ind)


def _write_spacing(style_el: etree._Element, prop: str, value: Any) -> None:
    """Merge into <w:spacing/>; before/after/line/lineRule from one prop at a time."""
    ppr = _get_or_create_ppr(style_el)
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = el("w:spacing")
        _ordered_insert(ppr, spacing, _PPR_CHILD_ORDER)
    if prop == "spacing_before":
        _set_or_clear_attr(spacing, "w:before", value)
    elif prop == "spacing_after":
        _set_or_clear_attr(spacing, "w:after", value)
    elif prop == "line_spacing":
        # Treat the rule as 'auto' unless one is already set on this element.
        # auto stores a multiplier (×240 → twips); exact/atLeast stores twips.
        if value is None:
            _set_or_clear_attr(spacing, "w:line", None)
            _set_or_clear_attr(spacing, "w:lineRule", None)
        else:
            existing_rule = spacing.get(qn("w:lineRule")) or "auto"
            if existing_rule == "auto":
                _set_or_clear_attr(spacing, "w:line", int(round(float(value) * 240)))
            else:
                _set_or_clear_attr(spacing, "w:line", int(round(float(value))))
            _set_or_clear_attr(spacing, "w:lineRule", existing_rule)
    elif prop == "line_spacing_rule":
        if value is None:
            _set_or_clear_attr(spacing, "w:lineRule", None)
        else:
            # Switching rules requires reinterpreting an existing line value;
            # we leave the numeric line value alone and let the caller re-set
            # line_spacing if they need a different scale.
            _set_or_clear_attr(spacing, "w:lineRule", value)
    if not spacing.attrib:
        ppr.remove(spacing)


def _write_font_name(style_el: etree._Element, value: str | None) -> None:
    """Write the four-typeface variant of <w:rFonts/>; clear all four on None."""
    rpr = _get_or_create_rpr(style_el)
    rfonts = rpr.find(qn("w:rFonts"))
    if value is None:
        if rfonts is not None:
            rpr.remove(rfonts)
        return
    if rfonts is None:
        rfonts = el("w:rFonts")
        _ordered_insert(rpr, rfonts, _RPR_CHILD_ORDER)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), value)


def _write_color(style_el: etree._Element, value: str | None) -> None:
    """Write a hex color. Accepts 'RRGGBB' (with/without leading #) or None."""
    if value is None:
        _remove_rpr_child(style_el, "color")
        return
    cleaned = value.lstrip("#").upper()
    if len(cleaned) != 6:
        raise ValueError(f"color_rgb expects RRGGBB hex, got {value!r}")
    int(cleaned, 16)  # validate hex
    _set_rpr_child(style_el, "color", {"w:val": cleaned})


# --------------------------------------------------------------------------
# Schema-ordered insert / replace helpers.
# --------------------------------------------------------------------------


def _ordered_insert(parent: etree._Element, child: etree._Element, order: tuple[str, ...]) -> None:
    """Insert ``child`` into ``parent`` at the position dictated by ``order``.

    Removes any existing element in ``parent`` with the same local name first,
    so calling twice with the same tag replaces the previous instance. The
    new element is placed immediately before the first existing sibling whose
    local name comes later in ``order``.
    """
    target_local = etree.QName(child.tag).localname
    for existing in parent.findall(qn(f"w:{target_local}")):
        parent.remove(existing)
    try:
        target_idx = order.index(target_local)
    except ValueError:
        parent.append(child)
        return
    for sibling in parent:
        if not isinstance(sibling.tag, str):
            continue
        sibling_local = etree.QName(sibling.tag).localname
        try:
            sibling_idx = order.index(sibling_local)
        except ValueError:
            continue
        if sibling_idx > target_idx:
            sibling.addprevious(child)
            return
    parent.append(child)


def _set_simple_child(style_el: etree._Element, tag: str, attrs: dict[str, str]) -> None:
    new_child = el(f"w:{tag}", **attrs)
    _ordered_insert(style_el, new_child, _STYLE_CHILD_ORDER)


def _get_or_create_ppr(style_el: etree._Element) -> etree._Element:
    ppr = style_el.find(qn("w:pPr"))
    if ppr is None:
        ppr = el("w:pPr")
        _ordered_insert(style_el, ppr, _STYLE_CHILD_ORDER)
    return ppr


def _get_or_create_rpr(style_el: etree._Element) -> etree._Element:
    rpr = style_el.find(qn("w:rPr"))
    if rpr is None:
        rpr = el("w:rPr")
        _ordered_insert(style_el, rpr, _STYLE_CHILD_ORDER)
    return rpr


def _set_ppr_child(style_el: etree._Element, tag: str, attrs: dict[str, str]) -> None:
    ppr = _get_or_create_ppr(style_el)
    new_child = el(f"w:{tag}", **attrs)
    _ordered_insert(ppr, new_child, _PPR_CHILD_ORDER)


def _set_rpr_child(style_el: etree._Element, tag: str, attrs: dict[str, str]) -> None:
    rpr = _get_or_create_rpr(style_el)
    new_child = el(f"w:{tag}", **attrs)
    _ordered_insert(rpr, new_child, _RPR_CHILD_ORDER)


def _remove_ppr_child(style_el: etree._Element, tag: str) -> None:
    ppr = style_el.find(qn("w:pPr"))
    if ppr is None:
        return
    for existing in ppr.findall(qn(f"w:{tag}")):
        ppr.remove(existing)
    if len(ppr) == 0:
        style_el.remove(ppr)


def _remove_rpr_child(style_el: etree._Element, tag: str) -> None:
    rpr = style_el.find(qn("w:rPr"))
    if rpr is None:
        return
    for existing in rpr.findall(qn(f"w:{tag}")):
        rpr.remove(existing)
    if len(rpr) == 0:
        style_el.remove(rpr)


def _remove_pr_child(style_el: etree._Element, container: Literal["ppr", "rpr"], tag: str) -> None:
    if container == "ppr":
        _remove_ppr_child(style_el, tag)
    else:
        _remove_rpr_child(style_el, tag)


def _set_or_clear_attr(node: etree._Element, key: str, value: Any) -> None:
    clark = qn(key) if ":" in key else key
    if value is None:
        if clark in node.attrib:
            del node.attrib[clark]
    else:
        node.set(clark, str(value))


# --------------------------------------------------------------------------
# apply_style helpers.
# --------------------------------------------------------------------------


def _set_paragraph_style(p_element: etree._Element, style_id: str) -> None:
    ppr = p_element.find(qn("w:pPr"))
    if ppr is None:
        ppr = el("w:pPr")
        p_element.insert(0, ppr)
    new_pstyle = el("w:pStyle", **{"w:val": style_id})
    for existing in ppr.findall(qn("w:pStyle")):
        ppr.remove(existing)
    ppr.insert(0, new_pstyle)


def _set_run_style(r_element: etree._Element, style_id: str) -> None:
    rpr = r_element.find(qn("w:rPr"))
    if rpr is None:
        rpr = el("w:rPr")
        r_element.insert(0, rpr)
    new_rstyle = el("w:rStyle", **{"w:val": style_id})
    for existing in rpr.findall(qn("w:rStyle")):
        rpr.remove(existing)
    rpr.insert(0, new_rstyle)


# --------------------------------------------------------------------------
# Lookup, reference scanning, list_styles helpers.
# --------------------------------------------------------------------------


def _find_style_element(styles_root: etree._Element, style_id: str) -> etree._Element | None:
    matches = xpath(styles_root, f"./w:style[@w:styleId='{style_id}']")
    return matches[0] if matches else None


def _find_references(doc: Document, style_id: str) -> list[etree._Element]:
    """Find references to ``style_id`` from the body and from other styles."""
    refs: list[etree._Element] = []
    body_root = doc.part.element
    # Body references: pStyle, rStyle, tblStyle.
    for tag in ("pStyle", "rStyle", "tblStyle"):
        refs.extend(xpath(body_root, f"//w:{tag}[@w:val='{style_id}']"))
    # Style-to-style references in styles.xml. Don't count the deletion target
    # itself even if it has a basedOn pointing nowhere — we want OUTBOUND
    # references *to* this id from other styles.
    styles_root = doc.styles.element
    for tag in ("basedOn", "next", "link", "numStyleLink", "styleLink"):
        for ref in xpath(styles_root, f"//w:{tag}[@w:val='{style_id}']"):
            owning_style = _enclosing_style_element(ref)
            if owning_style is None:
                continue
            if owning_style.get(qn("w:styleId")) == style_id:
                continue
            refs.append(ref)
    return refs


def _enclosing_style_element(node: etree._Element) -> etree._Element | None:
    cursor: etree._Element | None = node
    while cursor is not None:
        if isinstance(cursor.tag, str) and etree.QName(cursor.tag).localname == "style":
            return cursor
        cursor = cursor.getparent()
    return None


def _style_info_from_element(style_el: etree._Element) -> StyleInfo | None:
    sid = style_el.get(qn("w:styleId"))
    if sid is None:
        return None
    raw_type = style_el.get(qn("w:type")) or "paragraph"
    if raw_type not in ("paragraph", "character", "table", "numbering"):
        return None
    name_el = style_el.find(qn("w:name"))
    name = name_el.get(qn("w:val")) if name_el is not None else sid
    based = style_el.find(qn("w:basedOn"))
    based_id = based.get(qn("w:val")) if based is not None else None
    is_default = style_el.get(qn("w:default")) == "1"
    return StyleInfo(
        style_id=sid,
        name=name or sid,
        # raw_type is validated upstream against the StyleType literal set;
        # mypy can't see the narrowing through the caller's filter logic.
        style_type=raw_type,  # type: ignore[arg-type]
        based_on=based_id,
        is_default=is_default,
    )


# --------------------------------------------------------------------------
# Known built-ins table.
# --------------------------------------------------------------------------
# Definitions for the most common Word built-in styles. These match the
# structural shape Word writes (name, basedOn, link, uiPriority, qFormat) so
# Word treats them as the genuine built-ins rather than custom look-alikes.
# Property values for Heading 1-9 / Title / Subtitle / Quote / IntenseQuote
# are the post-Word-2013 "Office" defaults — sufficient for the resolver to
# return meaningful values; Word's UI will re-apply its own built-in styling
# regardless. ``custom`` is False for built-ins (they're not user-defined).
#
# Coverage: the styles SPEC §5 calls out as "at minimum". When extending,
# extract the structural fields from a freshly-Word-saved document rather
# than guessing — see IMPLEMENTATION.md §7 "Latent built-ins".

_BUILTIN_STYLES: dict[str, dict[str, Any]] = {
    "Normal": {
        "style_type": "paragraph",
        "name": "Normal",
        "ui_priority": 0,
        "q_format": True,
        "is_default": True,
    },
    "DefaultParagraphFont": {
        "style_type": "character",
        "name": "Default Paragraph Font",
        "ui_priority": 1,
        "is_default": True,
    },
    "TableNormal": {
        "style_type": "table",
        "name": "Normal Table",
        "ui_priority": 99,
        "is_default": True,
    },
    "NoList": {
        "style_type": "numbering",
        "name": "No List",
        "ui_priority": 99,
        "is_default": True,
    },
    "Heading1": {
        "style_type": "paragraph",
        "name": "heading 1",
        "based_on": "Normal",
        "next_style": "Normal",
        "linked_style": "Heading1Char",
        "ui_priority": 9,
        "q_format": True,
        "properties": {
            "keep_with_next": True,
            "keep_lines": True,
            "spacing_before": 240,
            "spacing_after": 0,
            "outline_level": 0,
            "font_size": 16.0,
            "color_rgb": "2F5496",
        },
    },
    "Heading2": {
        "style_type": "paragraph",
        "name": "heading 2",
        "based_on": "Normal",
        "next_style": "Normal",
        "linked_style": "Heading2Char",
        "ui_priority": 9,
        "q_format": True,
        "properties": {
            "keep_with_next": True,
            "keep_lines": True,
            "spacing_before": 40,
            "spacing_after": 0,
            "outline_level": 1,
            "font_size": 13.0,
            "color_rgb": "2F5496",
        },
    },
    "Heading3": {
        "style_type": "paragraph",
        "name": "heading 3",
        "based_on": "Normal",
        "next_style": "Normal",
        "linked_style": "Heading3Char",
        "ui_priority": 9,
        "q_format": True,
        "properties": {
            "keep_with_next": True,
            "keep_lines": True,
            "spacing_before": 40,
            "spacing_after": 0,
            "outline_level": 2,
            "font_size": 12.0,
            "color_rgb": "1F3763",
        },
    },
    "Heading4": {
        "style_type": "paragraph",
        "name": "heading 4",
        "based_on": "Normal",
        "next_style": "Normal",
        "linked_style": "Heading4Char",
        "ui_priority": 9,
        "q_format": True,
        "properties": {
            "keep_with_next": True,
            "keep_lines": True,
            "spacing_before": 40,
            "spacing_after": 0,
            "outline_level": 3,
            "italic": True,
            "color_rgb": "2F5496",
        },
    },
    "Heading5": {
        "style_type": "paragraph",
        "name": "heading 5",
        "based_on": "Normal",
        "next_style": "Normal",
        "linked_style": "Heading5Char",
        "ui_priority": 9,
        "q_format": True,
        "properties": {
            "keep_with_next": True,
            "keep_lines": True,
            "spacing_before": 40,
            "spacing_after": 0,
            "outline_level": 4,
            "color_rgb": "2F5496",
        },
    },
    "Heading6": {
        "style_type": "paragraph",
        "name": "heading 6",
        "based_on": "Normal",
        "next_style": "Normal",
        "linked_style": "Heading6Char",
        "ui_priority": 9,
        "q_format": True,
        "properties": {
            "keep_with_next": True,
            "keep_lines": True,
            "spacing_before": 40,
            "spacing_after": 0,
            "outline_level": 5,
            "color_rgb": "1F3763",
        },
    },
    "Heading7": {
        "style_type": "paragraph",
        "name": "heading 7",
        "based_on": "Normal",
        "next_style": "Normal",
        "linked_style": "Heading7Char",
        "ui_priority": 9,
        "q_format": True,
        "properties": {
            "keep_with_next": True,
            "keep_lines": True,
            "spacing_before": 40,
            "spacing_after": 0,
            "outline_level": 6,
            "italic": True,
            "color_rgb": "1F3763",
        },
    },
    "Heading8": {
        "style_type": "paragraph",
        "name": "heading 8",
        "based_on": "Normal",
        "next_style": "Normal",
        "linked_style": "Heading8Char",
        "ui_priority": 9,
        "q_format": True,
        "properties": {
            "keep_with_next": True,
            "keep_lines": True,
            "spacing_before": 40,
            "spacing_after": 0,
            "outline_level": 7,
            "font_size": 10.5,
            "color_rgb": "272727",
        },
    },
    "Heading9": {
        "style_type": "paragraph",
        "name": "heading 9",
        "based_on": "Normal",
        "next_style": "Normal",
        "linked_style": "Heading9Char",
        "ui_priority": 9,
        "q_format": True,
        "properties": {
            "keep_with_next": True,
            "keep_lines": True,
            "spacing_before": 40,
            "spacing_after": 0,
            "outline_level": 8,
            "italic": True,
            "font_size": 10.5,
            "color_rgb": "272727",
        },
    },
    "Title": {
        "style_type": "paragraph",
        "name": "Title",
        "based_on": "Normal",
        "next_style": "Normal",
        "linked_style": "TitleChar",
        "ui_priority": 10,
        "q_format": True,
        "properties": {
            "spacing_before": 0,
            "spacing_after": 80,
            "font_size": 28.0,
            "color_rgb": "000000",
        },
    },
    "Subtitle": {
        "style_type": "paragraph",
        "name": "Subtitle",
        "based_on": "Normal",
        "next_style": "Normal",
        "linked_style": "SubtitleChar",
        "ui_priority": 11,
        "q_format": True,
        "properties": {
            "spacing_after": 160,
            "font_size": 11.0,
            "italic": True,
            "color_rgb": "5A5A5A",
        },
    },
    "Quote": {
        "style_type": "paragraph",
        "name": "Quote",
        "based_on": "Normal",
        "next_style": "Normal",
        "linked_style": "QuoteChar",
        "ui_priority": 29,
        "q_format": True,
        "properties": {
            "italic": True,
            "color_rgb": "404040",
        },
    },
    "IntenseQuote": {
        "style_type": "paragraph",
        "name": "Intense Quote",
        "based_on": "Normal",
        "next_style": "Normal",
        "linked_style": "IntenseQuoteChar",
        "ui_priority": 30,
        "q_format": True,
        "properties": {
            "spacing_before": 360,
            "spacing_after": 360,
            "indent_left": 864,
            "indent_right": 864,
            "alignment": "center",
            "italic": True,
            "color_rgb": "0F4761",
        },
    },
    "ListParagraph": {
        "style_type": "paragraph",
        "name": "List Paragraph",
        "based_on": "Normal",
        "ui_priority": 34,
        "q_format": True,
        "properties": {
            "indent_left": 720,
        },
    },
    "Caption": {
        "style_type": "paragraph",
        "name": "caption",
        "based_on": "Normal",
        "next_style": "Normal",
        "ui_priority": 35,
        "q_format": True,
        "properties": {
            "spacing_after": 200,
            "line_spacing": 1.0,
            "font_size": 9.0,
            "italic": True,
            "color_rgb": "0E2841",
        },
    },
    "Hyperlink": {
        "style_type": "character",
        "name": "Hyperlink",
        "based_on": "DefaultParagraphFont",
        "ui_priority": 99,
        "properties": {
            "color_rgb": "0563C1",
            "underline": "single",
        },
    },
    "PlaceholderText": {
        "style_type": "character",
        "name": "Placeholder Text",
        "based_on": "DefaultParagraphFont",
        "ui_priority": 99,
        "properties": {
            "color_rgb": "808080",
        },
    },
    # ----------------------------------------------------------------------
    # Tier A — structural essentials (verified against python-docx's bundled
    # default.docx, where these are already materialised).
    # ----------------------------------------------------------------------
    "NoSpacing": {
        "style_type": "paragraph",
        "name": "No Spacing",
        "ui_priority": 1,
        "q_format": True,
        "properties": {
            "spacing_before": 0,
            "spacing_after": 0,
            "line_spacing": 1.0,
        },
    },
    "Header": {
        "style_type": "paragraph",
        "name": "header",
        "based_on": "Normal",
        "linked_style": "HeaderChar",
        "ui_priority": 99,
        "properties": {
            "spacing_after": 0,
            "line_spacing": 1.0,
        },
    },
    "HeaderChar": {
        "style_type": "character",
        "name": "Header Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "Header",
        "ui_priority": 99,
    },
    "Footer": {
        "style_type": "paragraph",
        "name": "footer",
        "based_on": "Normal",
        "linked_style": "FooterChar",
        "ui_priority": 99,
        "properties": {
            "spacing_after": 0,
            "line_spacing": 1.0,
        },
    },
    "FooterChar": {
        "style_type": "character",
        "name": "Footer Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "Footer",
        "ui_priority": 99,
    },
    "TableGrid": {
        "style_type": "table",
        "name": "Table Grid",
        "based_on": "TableNormal",
        "ui_priority": 59,
    },
    # ----------------------------------------------------------------------
    # Tier B — inline emphasis (character styles).
    # ----------------------------------------------------------------------
    "Strong": {
        "style_type": "character",
        "name": "Strong",
        "based_on": "DefaultParagraphFont",
        "ui_priority": 22,
        "q_format": True,
        "properties": {
            "bold": True,
        },
    },
    "Emphasis": {
        "style_type": "character",
        "name": "Emphasis",
        "based_on": "DefaultParagraphFont",
        "ui_priority": 20,
        "q_format": True,
        "properties": {
            "italic": True,
        },
    },
    "IntenseEmphasis": {
        "style_type": "character",
        "name": "Intense Emphasis",
        "based_on": "DefaultParagraphFont",
        "ui_priority": 21,
        "q_format": True,
        "properties": {
            "bold": True,
            "italic": True,
            "color_rgb": "4F81BD",
        },
    },
    "SubtleEmphasis": {
        "style_type": "character",
        "name": "Subtle Emphasis",
        "based_on": "DefaultParagraphFont",
        "ui_priority": 19,
        "q_format": True,
        "properties": {
            "italic": True,
            "color_rgb": "808080",
        },
    },
    "IntenseReference": {
        "style_type": "character",
        "name": "Intense Reference",
        "based_on": "DefaultParagraphFont",
        "ui_priority": 32,
        "q_format": True,
        "properties": {
            "bold": True,
            "small_caps": True,
            "color_rgb": "C0504D",
            "underline": "single",
        },
    },
    "SubtleReference": {
        "style_type": "character",
        "name": "Subtle Reference",
        "based_on": "DefaultParagraphFont",
        "ui_priority": 31,
        "q_format": True,
        "properties": {
            "small_caps": True,
            "color_rgb": "C0504D",
            "underline": "single",
        },
    },
    "BookTitle": {
        "style_type": "character",
        "name": "Book Title",
        "based_on": "DefaultParagraphFont",
        "ui_priority": 33,
        "q_format": True,
        "properties": {
            "bold": True,
            "small_caps": True,
        },
    },
    # ----------------------------------------------------------------------
    # Tier C — linked character styles for the heading / title family.
    # Word auto-creates these when the paragraph style has a w:link; carrying
    # them in the table lets ensure_style materialise them on docs that lack
    # them, so the link target isn't dangling.
    # ----------------------------------------------------------------------
    "Heading1Char": {
        "style_type": "character",
        "name": "Heading 1 Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "Heading1",
        "ui_priority": 9,
        "properties": {
            "bold": True,
            "font_size": 14.0,
            "color_rgb": "2F5496",
        },
    },
    "Heading2Char": {
        "style_type": "character",
        "name": "Heading 2 Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "Heading2",
        "ui_priority": 9,
        "properties": {
            "bold": True,
            "font_size": 13.0,
            "color_rgb": "2F5496",
        },
    },
    "Heading3Char": {
        "style_type": "character",
        "name": "Heading 3 Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "Heading3",
        "ui_priority": 9,
        "properties": {
            "bold": True,
            "font_size": 12.0,
            "color_rgb": "1F3763",
        },
    },
    "Heading4Char": {
        "style_type": "character",
        "name": "Heading 4 Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "Heading4",
        "ui_priority": 9,
        "properties": {
            "italic": True,
            "color_rgb": "2F5496",
        },
    },
    "Heading5Char": {
        "style_type": "character",
        "name": "Heading 5 Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "Heading5",
        "ui_priority": 9,
        "properties": {
            "color_rgb": "2F5496",
        },
    },
    "Heading6Char": {
        "style_type": "character",
        "name": "Heading 6 Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "Heading6",
        "ui_priority": 9,
        "properties": {
            "color_rgb": "1F3763",
        },
    },
    "Heading7Char": {
        "style_type": "character",
        "name": "Heading 7 Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "Heading7",
        "ui_priority": 9,
        "properties": {
            "italic": True,
            "color_rgb": "1F3763",
        },
    },
    "Heading8Char": {
        "style_type": "character",
        "name": "Heading 8 Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "Heading8",
        "ui_priority": 9,
        "properties": {
            "font_size": 10.5,
            "color_rgb": "272727",
        },
    },
    "Heading9Char": {
        "style_type": "character",
        "name": "Heading 9 Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "Heading9",
        "ui_priority": 9,
        "properties": {
            "italic": True,
            "font_size": 10.5,
            "color_rgb": "272727",
        },
    },
    "TitleChar": {
        "style_type": "character",
        "name": "Title Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "Title",
        "ui_priority": 10,
        "properties": {
            "font_size": 28.0,
            "color_rgb": "000000",
        },
    },
    "SubtitleChar": {
        "style_type": "character",
        "name": "Subtitle Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "Subtitle",
        "ui_priority": 11,
        "properties": {
            "italic": True,
            "font_size": 11.0,
            "color_rgb": "5A5A5A",
        },
    },
    "QuoteChar": {
        "style_type": "character",
        "name": "Quote Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "Quote",
        "ui_priority": 29,
        "properties": {
            "italic": True,
            "color_rgb": "404040",
        },
    },
    "IntenseQuoteChar": {
        "style_type": "character",
        "name": "Intense Quote Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "IntenseQuote",
        "ui_priority": 30,
        "properties": {
            "italic": True,
            "color_rgb": "0F4761",
        },
    },
    # ----------------------------------------------------------------------
    # Tier D — list paragraph variants (1–5 levels).
    # The numPr child Word writes for these styles is a placeholder without
    # @val (no concrete numbering link); we omit it here and rely on basedOn
    # plus indent. Callers wanting actual auto-numbering should attach a
    # numbering definition separately.
    # ----------------------------------------------------------------------
    "List": {
        "style_type": "paragraph",
        "name": "List",
        "based_on": "Normal",
        "ui_priority": 99,
        "properties": {
            "indent_left": 360,
            "indent_first_line": -360,
        },
    },
    "List2": {
        "style_type": "paragraph",
        "name": "List 2",
        "based_on": "Normal",
        "ui_priority": 99,
        "properties": {
            "indent_left": 720,
            "indent_first_line": -360,
        },
    },
    "List3": {
        "style_type": "paragraph",
        "name": "List 3",
        "based_on": "Normal",
        "ui_priority": 99,
        "properties": {
            "indent_left": 1080,
            "indent_first_line": -360,
        },
    },
    "ListBullet": {
        "style_type": "paragraph",
        "name": "List Bullet",
        "based_on": "Normal",
        "ui_priority": 99,
    },
    "ListBullet2": {
        "style_type": "paragraph",
        "name": "List Bullet 2",
        "based_on": "Normal",
        "ui_priority": 99,
    },
    "ListBullet3": {
        "style_type": "paragraph",
        "name": "List Bullet 3",
        "based_on": "Normal",
        "ui_priority": 99,
    },
    "ListBullet4": {
        "style_type": "paragraph",
        "name": "List Bullet 4",
        "based_on": "Normal",
        "ui_priority": 99,
    },
    "ListBullet5": {
        "style_type": "paragraph",
        "name": "List Bullet 5",
        "based_on": "Normal",
        "ui_priority": 99,
    },
    "ListNumber": {
        "style_type": "paragraph",
        "name": "List Number",
        "based_on": "Normal",
        "ui_priority": 99,
    },
    "ListNumber2": {
        "style_type": "paragraph",
        "name": "List Number 2",
        "based_on": "Normal",
        "ui_priority": 99,
    },
    "ListNumber3": {
        "style_type": "paragraph",
        "name": "List Number 3",
        "based_on": "Normal",
        "ui_priority": 99,
    },
    "ListNumber4": {
        "style_type": "paragraph",
        "name": "List Number 4",
        "based_on": "Normal",
        "ui_priority": 99,
    },
    "ListNumber5": {
        "style_type": "paragraph",
        "name": "List Number 5",
        "based_on": "Normal",
        "ui_priority": 99,
    },
    "ListContinue": {
        "style_type": "paragraph",
        "name": "List Continue",
        "based_on": "Normal",
        "ui_priority": 99,
        "properties": {
            "spacing_after": 120,
            "indent_left": 360,
        },
    },
    "ListContinue2": {
        "style_type": "paragraph",
        "name": "List Continue 2",
        "based_on": "Normal",
        "ui_priority": 99,
        "properties": {
            "spacing_after": 120,
            "indent_left": 720,
        },
    },
    "ListContinue3": {
        "style_type": "paragraph",
        "name": "List Continue 3",
        "based_on": "Normal",
        "ui_priority": 99,
        "properties": {
            "spacing_after": 120,
            "indent_left": 1080,
        },
    },
    "ListContinue4": {
        "style_type": "paragraph",
        "name": "List Continue 4",
        "based_on": "Normal",
        "ui_priority": 99,
        "properties": {
            "spacing_after": 120,
            "indent_left": 1440,
        },
    },
    "ListContinue5": {
        "style_type": "paragraph",
        "name": "List Continue 5",
        "based_on": "Normal",
        "ui_priority": 99,
        "properties": {
            "spacing_after": 120,
            "indent_left": 1800,
        },
    },
    # ----------------------------------------------------------------------
    # Tier E — TOC / index / table-of-* navigation styles.
    # TOC1–9 and TOC Heading defaults extracted from a Word-saved sample
    # (tests/fixtures/word_samples/sample-1.docx, 2026-05-19): TOC1 spacing_after=100
    # with no indent; TOC2..9 add a 240-twip progressive indent (240, 480, 720,
    # 960, 1200, 1440, 1680, 1920). TOCHeading is basedOn=Heading1 with a
    # 240-twip spacing_before and outline_level=9.
    # ----------------------------------------------------------------------
    "TOCHeading": {
        "style_type": "paragraph",
        "name": "TOC Heading",
        "based_on": "Heading1",
        "next_style": "Normal",
        "ui_priority": 39,
        "q_format": True,
        "properties": {
            "spacing_before": 240,
            "outline_level": 9,
        },
    },
    "TOC1": {
        "style_type": "paragraph",
        "name": "toc 1",
        "based_on": "Normal",
        "next_style": "Normal",
        "ui_priority": 39,
        "properties": {
            "spacing_after": 100,
        },
    },
    "TOC2": {
        "style_type": "paragraph",
        "name": "toc 2",
        "based_on": "Normal",
        "next_style": "Normal",
        "ui_priority": 39,
        "properties": {
            "spacing_after": 100,
            "indent_left": 240,
        },
    },
    "TOC3": {
        "style_type": "paragraph",
        "name": "toc 3",
        "based_on": "Normal",
        "next_style": "Normal",
        "ui_priority": 39,
        "properties": {
            "spacing_after": 100,
            "indent_left": 480,
        },
    },
    "TOC4": {
        "style_type": "paragraph",
        "name": "toc 4",
        "based_on": "Normal",
        "next_style": "Normal",
        "ui_priority": 39,
        "properties": {
            "spacing_after": 100,
            "indent_left": 720,
        },
    },
    "TOC5": {
        "style_type": "paragraph",
        "name": "toc 5",
        "based_on": "Normal",
        "next_style": "Normal",
        "ui_priority": 39,
        "properties": {
            "spacing_after": 100,
            "indent_left": 960,
        },
    },
    "TOC6": {
        "style_type": "paragraph",
        "name": "toc 6",
        "based_on": "Normal",
        "next_style": "Normal",
        "ui_priority": 39,
        "properties": {
            "spacing_after": 100,
            "indent_left": 1200,
        },
    },
    "TOC7": {
        "style_type": "paragraph",
        "name": "toc 7",
        "based_on": "Normal",
        "next_style": "Normal",
        "ui_priority": 39,
        "properties": {
            "spacing_after": 100,
            "indent_left": 1440,
        },
    },
    "TOC8": {
        "style_type": "paragraph",
        "name": "toc 8",
        "based_on": "Normal",
        "next_style": "Normal",
        "ui_priority": 39,
        "properties": {
            "spacing_after": 100,
            "indent_left": 1680,
        },
    },
    "TOC9": {
        "style_type": "paragraph",
        "name": "toc 9",
        "based_on": "Normal",
        "next_style": "Normal",
        "ui_priority": 39,
        "properties": {
            "spacing_after": 100,
            "indent_left": 1920,
        },
    },
    "IndexHeading": {
        "style_type": "paragraph",
        "name": "index heading",
        "based_on": "Normal",
        "next_style": "Index1",
        "ui_priority": 99,
        "properties": {
            "bold": True,
        },
    },
    "Index1": {
        "style_type": "paragraph",
        "name": "index 1",
        "based_on": "Normal",
        "next_style": "Normal",
        "ui_priority": 99,
        "properties": {
            "indent_left": 240,
            "indent_first_line": -240,
        },
    },
    # Word writes these with a lowercase 'o' in the styleId (TableofFigures,
    # not TableOfFigures) — verified against a Word-saved sample.
    "TableofFigures": {
        "style_type": "paragraph",
        "name": "table of figures",
        "based_on": "Normal",
        "next_style": "Normal",
        "ui_priority": 99,
        "properties": {
            "spacing_after": 0,
        },
    },
    "TableofAuthorities": {
        "style_type": "paragraph",
        "name": "table of authorities",
        "based_on": "Normal",
        "next_style": "Normal",
        "ui_priority": 99,
        "properties": {
            "spacing_after": 0,
            "indent_left": 240,
            "indent_first_line": -240,
        },
    },
    "TOAHeading": {
        "style_type": "paragraph",
        "name": "toa heading",
        "based_on": "Normal",
        "next_style": "Normal",
        "ui_priority": 99,
        "properties": {
            "spacing_before": 120,
            "bold": True,
        },
    },
    # ----------------------------------------------------------------------
    # Tier F — footnotes / endnotes / comments / balloons. Foot/endnote and
    # comment defaults extracted from sample-1.docx and Balloon* from
    # sample-2.docx (tests/fixtures/word_samples/, 2026-05-19): 10pt text,
    # single line spacing, superscript for the reference styles, Segoe UI
    # 9pt for Balloon*.
    # ----------------------------------------------------------------------
    "FootnoteText": {
        "style_type": "paragraph",
        "name": "footnote text",
        "based_on": "Normal",
        "linked_style": "FootnoteTextChar",
        "ui_priority": 99,
        "properties": {
            "spacing_after": 0,
            "font_size": 10.0,
        },
    },
    "FootnoteTextChar": {
        "style_type": "character",
        "name": "Footnote Text Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "FootnoteText",
        "ui_priority": 99,
        "properties": {
            "font_size": 10.0,
        },
    },
    "FootnoteReference": {
        "style_type": "character",
        "name": "footnote reference",
        "based_on": "DefaultParagraphFont",
        "ui_priority": 99,
        "properties": {
            "vert_align": "superscript",
        },
    },
    "EndnoteText": {
        "style_type": "paragraph",
        "name": "endnote text",
        "based_on": "Normal",
        "linked_style": "EndnoteTextChar",
        "ui_priority": 99,
        "properties": {
            "spacing_after": 0,
            "font_size": 10.0,
        },
    },
    "EndnoteTextChar": {
        "style_type": "character",
        "name": "Endnote Text Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "EndnoteText",
        "ui_priority": 99,
        "properties": {
            "font_size": 10.0,
        },
    },
    "EndnoteReference": {
        "style_type": "character",
        "name": "endnote reference",
        "based_on": "DefaultParagraphFont",
        "ui_priority": 99,
        "properties": {
            "vert_align": "superscript",
        },
    },
    "CommentText": {
        "style_type": "paragraph",
        "name": "annotation text",
        "based_on": "Normal",
        "linked_style": "CommentTextChar",
        "ui_priority": 99,
        "properties": {
            "line_spacing": 1.0,
            "font_size": 10.0,
        },
    },
    "CommentTextChar": {
        "style_type": "character",
        "name": "Comment Text Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "CommentText",
        "ui_priority": 99,
        "properties": {
            "font_size": 10.0,
        },
    },
    "CommentReference": {
        "style_type": "character",
        "name": "annotation reference",
        "based_on": "DefaultParagraphFont",
        "ui_priority": 99,
        "properties": {
            "font_size": 8.0,
        },
    },
    "CommentSubject": {
        "style_type": "paragraph",
        "name": "annotation subject",
        "based_on": "CommentText",
        "next_style": "CommentText",
        "linked_style": "CommentSubjectChar",
        "ui_priority": 99,
        "properties": {
            "bold": True,
        },
    },
    "CommentSubjectChar": {
        "style_type": "character",
        "name": "Comment Subject Char",
        "based_on": "CommentTextChar",
        "linked_style": "CommentSubject",
        "ui_priority": 99,
        "properties": {
            "bold": True,
            "font_size": 10.0,
        },
    },
    "BalloonText": {
        "style_type": "paragraph",
        "name": "Balloon Text",
        "based_on": "Normal",
        "linked_style": "BalloonTextChar",
        "ui_priority": 99,
        "properties": {
            "font_name": "Segoe UI",
            "font_size": 9.0,
            "spacing_after": 0,
            "line_spacing": 1.0,
        },
    },
    "BalloonTextChar": {
        "style_type": "character",
        "name": "Balloon Text Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "BalloonText",
        "ui_priority": 99,
        "properties": {
            "font_name": "Segoe UI",
            "font_size": 9.0,
        },
    },
    # ----------------------------------------------------------------------
    # Tier G — misc text-block styles (code, macro, indents). BodyText/2/3
    # and their Char companions, MacroText/Char, and the Header/Footer pair
    # (Tier A above) all extracted from a Word-saved sample (2026-05-19).
    # Tabs on Header/Footer/MacroText and theme fonts on IndexHeading are
    # known limitations of the property writer and intentionally omitted.
    # ----------------------------------------------------------------------
    "BodyText": {
        "style_type": "paragraph",
        "name": "Body Text",
        "based_on": "Normal",
        "linked_style": "BodyTextChar",
        "ui_priority": 99,
        "properties": {
            "spacing_after": 120,
        },
    },
    "BodyTextChar": {
        "style_type": "character",
        "name": "Body Text Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "BodyText",
        "ui_priority": 99,
    },
    "BodyText2": {
        "style_type": "paragraph",
        "name": "Body Text 2",
        "based_on": "Normal",
        "linked_style": "BodyText2Char",
        "ui_priority": 99,
        "properties": {
            "spacing_after": 120,
            "line_spacing": 2.0,
        },
    },
    "BodyText2Char": {
        "style_type": "character",
        "name": "Body Text 2 Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "BodyText2",
        "ui_priority": 99,
    },
    "BodyText3": {
        "style_type": "paragraph",
        "name": "Body Text 3",
        "based_on": "Normal",
        "linked_style": "BodyText3Char",
        "ui_priority": 99,
        "properties": {
            "spacing_after": 120,
            "font_size": 8.0,
        },
    },
    "BodyText3Char": {
        "style_type": "character",
        "name": "Body Text 3 Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "BodyText3",
        "ui_priority": 99,
        "properties": {
            "font_size": 8.0,
        },
    },
    "MacroText": {
        "style_type": "paragraph",
        "name": "macro",
        "linked_style": "MacroTextChar",
        "ui_priority": 99,
        "properties": {
            "spacing_after": 0,
            "font_name": "Consolas",
            "font_size": 10.0,
        },
    },
    "MacroTextChar": {
        "style_type": "character",
        "name": "Macro Text Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "MacroText",
        "ui_priority": 99,
        "properties": {
            "font_name": "Consolas",
            "font_size": 10.0,
        },
    },
    "HTMLPreformatted": {
        "style_type": "paragraph",
        "name": "HTML Preformatted",
        "based_on": "Normal",
        "linked_style": "HTMLPreformattedChar",
        "ui_priority": 99,
        "properties": {
            "font_name": "Consolas",
            "font_size": 10.0,
        },
    },
    "HTMLPreformattedChar": {
        "style_type": "character",
        "name": "HTML Preformatted Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "HTMLPreformatted",
        "ui_priority": 99,
        "properties": {
            "font_name": "Consolas",
            "font_size": 10.0,
        },
    },
    "PlainText": {
        "style_type": "paragraph",
        "name": "Plain Text",
        "based_on": "Normal",
        "linked_style": "PlainTextChar",
        "ui_priority": 99,
        "properties": {
            "font_name": "Consolas",
            "font_size": 10.5,
        },
    },
    "PlainTextChar": {
        "style_type": "character",
        "name": "Plain Text Char",
        "based_on": "DefaultParagraphFont",
        "linked_style": "PlainText",
        "ui_priority": 99,
        "properties": {
            "font_name": "Consolas",
            "font_size": 10.5,
        },
    },
    "NormalIndent": {
        "style_type": "paragraph",
        "name": "Normal Indent",
        "based_on": "Normal",
        "ui_priority": 99,
        "properties": {
            "indent_left": 720,
        },
    },
    "BlockText": {
        "style_type": "paragraph",
        "name": "Block Text",
        "based_on": "Normal",
        "ui_priority": 99,
        "properties": {
            "indent_left": 1152,
            "indent_right": 1152,
            "italic": True,
            "color_rgb": "156082",
        },
    },
}


def _materialise_builtin(doc: Document, style_id: str, spec: dict[str, Any]) -> StyleProxy:
    """Create a style from a known-built-ins table entry. Internal."""
    properties = spec.get("properties", {})
    proxy = create_style(
        doc,
        style_id,
        style_type=spec.get("style_type", "paragraph"),
        name=spec.get("name"),
        based_on=spec.get("based_on"),
        next_style=spec.get("next_style"),
        linked_style=spec.get("linked_style"),
        ui_priority=spec.get("ui_priority", 99),
        q_format=spec.get("q_format", False),
        custom=False,
        **properties,
    )
    if spec.get("is_default"):
        proxy.element.set(qn("w:default"), "1")
    return proxy


__all__ = [
    "StyleExistsError",
    "StyleInUseError",
    "StyleInfo",
    "StyleNotFoundError",
    "StyleProxy",
    "UnknownStylePropertyError",
    "apply_style",
    "create_style",
    "delete_style",
    "ensure_style",
    "find_matching_style",
    "list_styles",
    "modify_style",
    "remap_styles",
]
