"""Cascade resolver: ``resolve_effective_formatting``.

Walks the six layers of OOXML formatting precedence (SPEC §4) and returns a
fully-resolved :class:`ResolvedFormatting` describing what a paragraph, run,
or cell would render with right now. Later layers override earlier ones,
except toggle properties (bold, italic, etc.) which XOR through the chain
per ECMA-376 17.7.3.

Provenance tracking is plumbed through the same walk gated by the
``include_provenance`` flag; with the flag off, the resolver's value output
is identical (verified by ``test_provenance_does_not_change_values``).
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from typing import TYPE_CHECKING, Any, Literal

from lxml import etree

from docx_plus.core import DocxPlusError
from docx_plus.core.ns import qn
from docx_plus.core.oxml import xpath
from docx_plus.styles.theme import ThemeColors, load_theme, resolve_theme_color

if TYPE_CHECKING:
    from docx.document import Document
    from docx.table import _Cell
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run


_MAX_STYLE_CHAIN_DEPTH = 11

# Toggle properties XOR through the cascade per ECMA-376 17.7.3.
# Mapped from rPr child name to ResolvedFormatting field name.
_TOGGLE_RPR: dict[str, str] = {
    "b": "bold",
    "i": "italic",
    "caps": "caps",
    "smallCaps": "small_caps",
    "strike": "strike",
    "vanish": "vanish",
}
# Other toggles per the spec (bCs, iCs, emboss, imprint, outline, shadow,
# dstrike notably NOT a toggle) are accepted but not yet surfaced on the
# ResolvedFormatting dataclass; v0.2 may expand it.


Layer = Literal[
    "docDefaults",
    "tableStyle",
    "paragraphStyle",
    "linkedCharStyle",
    "numbering",
    "directParagraph",
    "directRun",
]


class StyleCascadeError(DocxPlusError):
    """Raised when the basedOn chain cycles or exceeds Word's depth limit."""


class MissingPartError(DocxPlusError):
    """Raised when a referenced document part (e.g. numbering.xml) is absent."""


@dataclass(frozen=True)
class FormattingSource:
    """Identifies the cascade layer that contributed a resolved property.

    ``layer`` is the cascade layer the value came from. For style layers,
    ``style_id`` names the specific style (the lowest one in the basedOn
    chain that set the value); ``chain_depth`` records how many basedOn hops
    away that style was from the target. ``is_toggle_resolved`` is True when
    the value is the XOR result across multiple layers rather than a direct
    set.
    """

    layer: Layer
    style_id: str | None = None
    is_toggle_resolved: bool = False
    chain_depth: int | None = None


@dataclass(frozen=True)
class ResolvedFormatting:
    """The effective formatting for a paragraph, run, or table cell.

    Every field is ``None`` until some layer of the cascade sets it. Toggle
    properties carry their XOR-resolved boolean. SPEC §4 specifies the fields.
    """

    # Identity
    style_id: str | None = None
    style_name: str | None = None

    # Paragraph-level
    alignment: str | None = None
    indent_left: int | None = None
    indent_right: int | None = None
    indent_first_line: int | None = None
    spacing_before: int | None = None
    spacing_after: int | None = None
    line_spacing: float | None = None
    line_spacing_rule: str | None = None
    keep_with_next: bool | None = None
    keep_lines: bool | None = None
    page_break_before: bool | None = None
    outline_level: int | None = None

    # Run-level
    font_name: str | None = None
    font_size: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: str | None = None
    strike: bool | None = None
    color_rgb: str | None = None
    highlight: str | None = None
    caps: bool | None = None
    small_caps: bool | None = None
    vanish: bool | None = None
    vert_align: str | None = None

    # Numbering
    num_id: int | None = None
    num_level: int | None = None

    # Meta
    partial: bool = False
    provenance: dict[str, FormattingSource] | None = None


def resolve_effective_formatting(
    target: Paragraph | Run | _Cell,
    *,
    include_provenance: bool = False,
) -> ResolvedFormatting:
    """Resolve the effective formatting for ``target``.

    Walks the six cascade layers in precedence order, returning a fully
    resolved :class:`ResolvedFormatting`. Toggle properties XOR through the
    chain per ECMA-376 17.7.3. Theme colors are resolved against the
    document's theme part; if the theme is missing or malformed, the result's
    ``partial`` flag is set and unresolved theme names are returned in place
    of hex values.

    Args:
        target: A python-docx :class:`~docx.text.paragraph.Paragraph`,
            :class:`~docx.text.run.Run`, or :class:`~docx.table._Cell`.
        include_provenance: If True, populate ``.provenance`` with the cascade
            layer that set each field. Default False.

    Returns:
        A :class:`ResolvedFormatting` snapshot.

    Raises:
        StyleCascadeError: If the basedOn chain has a cycle or exceeds Word's
            depth limit of 11.
        MissingPartError: If the target's paragraph references a numbering id
            that exists but the ``numbering.xml`` part itself is absent.

    Example:
        >>> from docx import Document
        >>> from docx_plus.styles.inspect import resolve_effective_formatting
        >>> doc = Document()
        >>> p = doc.add_paragraph("Hello")
        >>> resolved = resolve_effective_formatting(p)
        >>> resolved.font_size  # e.g. 11.0 from docDefaults
        11.0
    """
    target_kind = _classify_target(target)
    doc = _document_of(target)
    styles_root = doc.styles.element
    theme = load_theme(doc)

    acc = _Accumulator(theme=theme, want_provenance=include_provenance)
    if theme is None:
        acc.partial = True

    if target_kind == "paragraph":
        _apply_paragraph_cascade(acc, doc, styles_root, target._p)  # type: ignore[union-attr]
    elif target_kind == "run":
        paragraph_element = _enclosing_paragraph(target._r)  # type: ignore[union-attr]
        _apply_paragraph_cascade(
            acc,
            doc,
            styles_root,
            paragraph_element,
            run_element=target._r,  # type: ignore[union-attr]
        )
    else:  # cell
        _apply_cell_cascade(acc, doc, styles_root, target._tc)  # type: ignore[union-attr]

    return acc.freeze()


# --------------------------------------------------------------------------
# Accumulator: in-progress resolved state, with optional provenance tracking.
# --------------------------------------------------------------------------


@dataclass
class _Accumulator:
    """Mutable in-progress state during the cascade walk."""

    theme: ThemeColors | None
    want_provenance: bool
    values: dict[str, Any] = field(default_factory=dict)
    provenance: dict[str, FormattingSource] = field(default_factory=dict)
    partial: bool = False

    def set(self, name: str, value: Any, source: FormattingSource) -> None:
        """Set a non-toggle property, recording provenance if requested."""
        if value is None:
            return
        self.values[name] = value
        if self.want_provenance:
            self.provenance[name] = source

    def toggle(self, name: str, val_attr: str | None, source: FormattingSource) -> None:
        """Apply XOR toggle rule to ``name`` per ECMA-376 17.7.3.

        ``val_attr`` is the ``w:val`` attribute on the toggle element, or
        ``None`` if absent.
        """
        if val_attr in ("0", "false"):
            new_value = False
            toggle_resolved = False
        else:
            current = self.values.get(name)
            new_value = True if current is None else not current
            toggle_resolved = current is not None
        self.values[name] = new_value
        if self.want_provenance:
            self.provenance[name] = replace(source, is_toggle_resolved=toggle_resolved)

    def freeze(self) -> ResolvedFormatting:
        """Snapshot into an immutable :class:`ResolvedFormatting`."""
        kwargs: dict[str, Any] = dict(self.values)
        kwargs["partial"] = self.partial
        kwargs["provenance"] = dict(self.provenance) if self.want_provenance else None
        return ResolvedFormatting(**kwargs)


# --------------------------------------------------------------------------
# Cascade entry points.
# --------------------------------------------------------------------------


def _apply_paragraph_cascade(
    acc: _Accumulator,
    doc: Document,
    styles_root: etree._Element,
    p_element: etree._Element,
    run_element: etree._Element | None = None,
) -> None:
    """Walk layers 1, 3, 4, 5 (and 6 if run_element) for a paragraph target."""
    # Layer 1: docDefaults
    _apply_doc_defaults(acc, styles_root)

    # Layer 2: table style (if inside a table)
    enclosing_tc = _enclosing_cell(p_element)
    if enclosing_tc is not None:
        table_element = _enclosing_table(enclosing_tc)
        if table_element is not None:
            _apply_table_style_chain(acc, styles_root, table_element)

    # Layer 3: paragraph style chain
    p_style_id = _paragraph_style_id(p_element)
    if p_style_id is not None:
        acc.set(
            "style_id",
            p_style_id,
            FormattingSource(layer="paragraphStyle", style_id=p_style_id, chain_depth=0),
        )
        style_name = _style_name(styles_root, p_style_id)
        if style_name is not None:
            acc.set(
                "style_name",
                style_name,
                FormattingSource(layer="paragraphStyle", style_id=p_style_id, chain_depth=0),
            )
        _apply_style_chain(acc, styles_root, p_style_id, "paragraphStyle")

    # Layer 4: numbering
    num_pr = p_element.find(f"./{qn('w:pPr')}/{qn('w:numPr')}")
    if num_pr is not None:
        _apply_numbering(acc, doc, num_pr)

    # Layer 5: direct paragraph formatting
    direct_ppr = p_element.find(qn("w:pPr"))
    if direct_ppr is not None:
        _apply_ppr(acc, direct_ppr, FormattingSource(layer="directParagraph"))
        # rPr inside pPr (paragraph mark formatting) — affects whole-paragraph runs
        direct_ppr_rpr = direct_ppr.find(qn("w:rPr"))
        if direct_ppr_rpr is not None and run_element is None:
            _apply_rpr(acc, direct_ppr_rpr, FormattingSource(layer="directParagraph"))

    if run_element is not None:
        # Linked character style (for Run targets only), per SPEC §4.
        if p_style_id is not None:
            linked_id = _linked_style_id(styles_root, p_style_id)
            if linked_id is not None:
                _apply_style_chain(acc, styles_root, linked_id, "linkedCharStyle")

        # Layer 6: direct run formatting
        run_rpr = run_element.find(qn("w:rPr"))
        if run_rpr is not None:
            _apply_rpr(acc, run_rpr, FormattingSource(layer="directRun"))
        # Run-level rStyle reference (character style applied to one run)
        run_style_id = _run_style_id(run_element)
        if run_style_id is not None:
            _apply_style_chain(acc, styles_root, run_style_id, "linkedCharStyle")


def _apply_cell_cascade(
    acc: _Accumulator,
    doc: Document,  # noqa: ARG001
    styles_root: etree._Element,
    tc_element: etree._Element,
) -> None:
    """Resolve formatting for a table cell — table style chain only, for now."""
    _apply_doc_defaults(acc, styles_root)
    table_element = _enclosing_table(tc_element)
    if table_element is not None:
        _apply_table_style_chain(acc, styles_root, table_element)


# --------------------------------------------------------------------------
# Layer helpers.
# --------------------------------------------------------------------------


def _apply_doc_defaults(acc: _Accumulator, styles_root: etree._Element) -> None:
    defaults = styles_root.find(qn("w:docDefaults"))
    if defaults is None:
        return
    source = FormattingSource(layer="docDefaults")

    rpr_default = defaults.find(qn("w:rPrDefault"))
    if rpr_default is not None:
        rpr = rpr_default.find(qn("w:rPr"))
        if rpr is not None:
            _apply_rpr(acc, rpr, source)

    ppr_default = defaults.find(qn("w:pPrDefault"))
    if ppr_default is not None:
        ppr = ppr_default.find(qn("w:pPr"))
        if ppr is not None:
            _apply_ppr(acc, ppr, source)


def _apply_style_chain(
    acc: _Accumulator,
    styles_root: etree._Element,
    leaf_style_id: str,
    layer: Layer,
) -> None:
    """Walk the basedOn chain and apply each style's pPr/rPr ancestors-first."""
    chain = _collect_style_chain(styles_root, leaf_style_id)
    # Apply in reverse: deepest ancestor first so leaf (most specific) wins.
    for depth, (style_id, style_el) in enumerate(reversed(chain)):
        chain_depth = len(chain) - 1 - depth
        source = FormattingSource(layer=layer, style_id=style_id, chain_depth=chain_depth)
        ppr = style_el.find(qn("w:pPr"))
        if ppr is not None:
            _apply_ppr(acc, ppr, source)
        rpr = style_el.find(qn("w:rPr"))
        if rpr is not None:
            _apply_rpr(acc, rpr, source)


def _collect_style_chain(
    styles_root: etree._Element, leaf_style_id: str
) -> list[tuple[str, etree._Element]]:
    """Return [(id, element), ...] from leaf to root, with cycle/depth checks."""
    chain: list[tuple[str, etree._Element]] = []
    visited: set[str] = set()
    current_id: str | None = leaf_style_id
    while current_id is not None:
        if current_id in visited:
            cycle_path = " -> ".join([sid for sid, _ in chain] + [current_id])
            raise StyleCascadeError(f"cycle in basedOn chain: {cycle_path}")
        if len(chain) > _MAX_STYLE_CHAIN_DEPTH:
            chain_ids = " -> ".join(sid for sid, _ in chain)
            raise StyleCascadeError(
                f"basedOn chain exceeds depth {_MAX_STYLE_CHAIN_DEPTH}: {chain_ids}"
            )
        style_el = _find_style(styles_root, current_id)
        if style_el is None:
            break
        chain.append((current_id, style_el))
        visited.add(current_id)
        based_on = style_el.find(qn("w:basedOn"))
        current_id = based_on.get(qn("w:val")) if based_on is not None else None
    return chain


def _apply_table_style_chain(
    acc: _Accumulator,
    styles_root: etree._Element,
    tbl_element: etree._Element,
) -> None:
    """Apply the table's style chain — base only; conditional formatting deferred.

    Only the base pPr/rPr from each style in the basedOn chain is applied.
    Conditional formatting via ``w:tblStylePr`` (firstRow, lastRow, firstCol,
    etc.) is recognised in the spec but deferred — see SPEC §4 step 2.
    """
    tbl_pr = tbl_element.find(qn("w:tblPr"))
    if tbl_pr is None:
        return
    tbl_style = tbl_pr.find(qn("w:tblStyle"))
    if tbl_style is None:
        return
    style_id = tbl_style.get(qn("w:val"))
    if style_id is None:
        return
    _apply_style_chain(acc, styles_root, style_id, "tableStyle")


def _apply_numbering(acc: _Accumulator, doc: Document, num_pr: etree._Element) -> None:
    num_id_el = num_pr.find(qn("w:numId"))
    ilvl_el = num_pr.find(qn("w:ilvl"))
    if num_id_el is None:
        return
    num_id_raw = num_id_el.get(qn("w:val"))
    if num_id_raw is None:
        return
    try:
        num_id = int(num_id_raw)
    except ValueError:
        return
    ilvl = 0
    if ilvl_el is not None:
        ilvl_raw = ilvl_el.get(qn("w:val"))
        if ilvl_raw is not None:
            try:
                ilvl = int(ilvl_raw)
            except ValueError:
                ilvl = 0
    source = FormattingSource(layer="numbering")
    acc.set("num_id", num_id, source)
    acc.set("num_level", ilvl, source)

    numbering_root = _numbering_root(doc)
    if numbering_root is None:
        # numPr references a num that isn't materialised — common when Word
        # hasn't authored the numbering part yet. Not fatal.
        return
    abstract_num = _resolve_abstract_num(numbering_root, num_id)
    if abstract_num is None:
        return
    lvl_el = _find_level(abstract_num, ilvl)
    if lvl_el is None:
        return
    lvl_ppr = lvl_el.find(qn("w:pPr"))
    if lvl_ppr is not None:
        _apply_ppr(acc, lvl_ppr, source)
    lvl_rpr = lvl_el.find(qn("w:rPr"))
    if lvl_rpr is not None:
        _apply_rpr(acc, lvl_rpr, source)


# --------------------------------------------------------------------------
# pPr / rPr property extraction.
# --------------------------------------------------------------------------


def _apply_ppr(acc: _Accumulator, ppr: etree._Element, source: FormattingSource) -> None:
    jc = ppr.find(qn("w:jc"))
    if jc is not None:
        acc.set("alignment", jc.get(qn("w:val")), source)

    ind = ppr.find(qn("w:ind"))
    if ind is not None:
        _apply_indent(acc, ind, source)

    spacing = ppr.find(qn("w:spacing"))
    if spacing is not None:
        _apply_spacing(acc, spacing, source)

    for tag, field_name in (
        ("keepNext", "keep_with_next"),
        ("keepLines", "keep_lines"),
        ("pageBreakBefore", "page_break_before"),
    ):
        flag_el = ppr.find(qn(f"w:{tag}"))
        if flag_el is not None:
            raw = flag_el.get(qn("w:val"))
            acc.set(field_name, raw not in ("0", "false"), source)

    outline = ppr.find(qn("w:outlineLvl"))
    if outline is not None:
        raw = outline.get(qn("w:val"))
        if raw is not None:
            try:
                acc.set("outline_level", int(raw), source)
            except ValueError:
                pass


def _apply_indent(acc: _Accumulator, ind: etree._Element, source: FormattingSource) -> None:
    left = ind.get(qn("w:left")) or ind.get(qn("w:start"))
    right = ind.get(qn("w:right")) or ind.get(qn("w:end"))
    first_line = ind.get(qn("w:firstLine"))
    hanging = ind.get(qn("w:hanging"))
    if left is not None:
        try:
            acc.set("indent_left", int(left), source)
        except ValueError:
            pass
    if right is not None:
        try:
            acc.set("indent_right", int(right), source)
        except ValueError:
            pass
    if hanging is not None:
        try:
            acc.set("indent_first_line", -int(hanging), source)
        except ValueError:
            pass
    elif first_line is not None:
        try:
            acc.set("indent_first_line", int(first_line), source)
        except ValueError:
            pass


def _apply_spacing(acc: _Accumulator, spacing: etree._Element, source: FormattingSource) -> None:
    before = spacing.get(qn("w:before"))
    after = spacing.get(qn("w:after"))
    line = spacing.get(qn("w:line"))
    line_rule = spacing.get(qn("w:lineRule"))
    if before is not None:
        try:
            acc.set("spacing_before", int(before), source)
        except ValueError:
            pass
    if after is not None:
        try:
            acc.set("spacing_after", int(after), source)
        except ValueError:
            pass
    if line is not None:
        try:
            line_val = int(line)
        except ValueError:
            return
        rule = line_rule or "auto"
        if rule == "auto":
            acc.set("line_spacing", line_val / 240.0, source)
        else:
            acc.set("line_spacing", float(line_val), source)
        acc.set("line_spacing_rule", rule, source)


def _apply_rpr(acc: _Accumulator, rpr: etree._Element, source: FormattingSource) -> None:
    for child in rpr:
        if not isinstance(child.tag, str):
            continue
        local = etree.QName(child.tag).localname
        if local in _TOGGLE_RPR:
            field_name = _TOGGLE_RPR[local]
            acc.toggle(field_name, child.get(qn("w:val")), source)
            continue
        if local == "rFonts":
            font = (
                child.get(qn("w:asciiTheme"))
                or child.get(qn("w:ascii"))
                or child.get(qn("w:hAnsi"))
                or child.get(qn("w:cs"))
            )
            if font is not None:
                resolved = _resolve_font_theme(font, source)
                acc.set("font_name", resolved, source)
        elif local == "sz":
            raw = child.get(qn("w:val"))
            if raw is not None:
                try:
                    acc.set("font_size", int(raw) / 2.0, source)
                except ValueError:
                    pass
        elif local == "color":
            color_val = _resolve_color(child, acc)
            if color_val is not None:
                acc.set("color_rgb", color_val, source)
        elif local == "u":
            val = child.get(qn("w:val"))
            if val is not None:
                acc.set("underline", val, source)
        elif local == "highlight":
            val = child.get(qn("w:val"))
            if val is not None:
                acc.set("highlight", val, source)
        elif local == "vertAlign":
            val = child.get(qn("w:val"))
            if val is not None:
                acc.set("vert_align", val, source)


def _resolve_color(color_el: etree._Element, acc: _Accumulator) -> str | None:
    theme_name = color_el.get(qn("w:themeColor"))
    if theme_name is not None:
        tint = color_el.get(qn("w:themeTint"))
        shade = color_el.get(qn("w:themeShade"))
        resolved = resolve_theme_color(acc.theme, theme_name, tint=tint, shade=shade)
        if resolved is not None:
            return resolved
        # Theme requested but theme not resolvable — emit unresolved name and
        # mark the result partial. SPEC §4 "Theme resolution edge cases".
        acc.partial = True
        return theme_name
    val = color_el.get(qn("w:val"))
    if val and val.lower() != "auto":
        return val.upper()
    return None


def _resolve_font_theme(value: str, source: FormattingSource) -> str:  # noqa: ARG001
    """Theme font tokens (majorAscii etc.) pass through as-is for v0.1.

    Theme font resolution would require reading ``a:fontScheme`` from the
    theme part; deferred until a caller actually depends on the resolved
    typeface name. Returning the token preserves enough information for
    diagnostic output.
    """
    return value


# --------------------------------------------------------------------------
# Document traversal / metadata helpers.
# --------------------------------------------------------------------------


def _classify_target(target: object) -> Literal["paragraph", "run", "cell"]:
    from docx.table import _Cell
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run

    if isinstance(target, Paragraph):
        return "paragraph"
    if isinstance(target, Run):
        return "run"
    if isinstance(target, _Cell):
        return "cell"
    kind = type(target).__name__
    raise TypeError(f"resolve_effective_formatting expects Paragraph, Run, or _Cell; got {kind}")


def _document_of(target: Paragraph | Run | _Cell) -> Document:
    """Return the owning ``Document`` for a paragraph/run/cell.

    python-docx exposes ``.document`` on the main document part. We cast
    through ``Any`` because the base ``Part`` class in python-docx is not
    typed with that attribute even though concrete subclasses provide it.
    """
    part: Any = target.part
    doc: Document = part.document
    return doc


def _find_style(styles_root: etree._Element, style_id: str) -> etree._Element | None:
    matches = xpath(styles_root, f"./w:style[@w:styleId='{style_id}']")
    return matches[0] if matches else None


def _style_name(styles_root: etree._Element, style_id: str) -> str | None:
    style_el = _find_style(styles_root, style_id)
    if style_el is None:
        return None
    name_el = style_el.find(qn("w:name"))
    if name_el is None:
        return None
    return name_el.get(qn("w:val"))


def _linked_style_id(styles_root: etree._Element, style_id: str) -> str | None:
    style_el = _find_style(styles_root, style_id)
    if style_el is None:
        return None
    link_el = style_el.find(qn("w:link"))
    if link_el is None:
        return None
    return link_el.get(qn("w:val"))


def _paragraph_style_id(p_element: etree._Element) -> str | None:
    pstyle = p_element.find(f"./{qn('w:pPr')}/{qn('w:pStyle')}")
    if pstyle is None:
        return None
    return pstyle.get(qn("w:val"))


def _run_style_id(r_element: etree._Element) -> str | None:
    rstyle = r_element.find(f"./{qn('w:rPr')}/{qn('w:rStyle')}")
    if rstyle is None:
        return None
    return rstyle.get(qn("w:val"))


def _enclosing_paragraph(r_element: etree._Element) -> etree._Element:
    node: etree._Element | None = r_element
    while node is not None:
        if isinstance(node.tag, str) and etree.QName(node.tag).localname == "p":
            return node
        node = node.getparent()
    raise ValueError("run is not inside a paragraph")


def _enclosing_cell(p_element: etree._Element) -> etree._Element | None:
    node: etree._Element | None = p_element.getparent()
    while node is not None:
        if isinstance(node.tag, str) and etree.QName(node.tag).localname == "tc":
            return node
        node = node.getparent()
    return None


def _enclosing_table(node: etree._Element) -> etree._Element | None:
    cursor: etree._Element | None = node
    while cursor is not None:
        if isinstance(cursor.tag, str) and etree.QName(cursor.tag).localname == "tbl":
            return cursor
        cursor = cursor.getparent()
    return None


# --------------------------------------------------------------------------
# Numbering helpers.
# --------------------------------------------------------------------------


def _numbering_root(doc: Document) -> etree._Element | None:
    numbering_part = getattr(doc.part, "numbering_part", None)
    if numbering_part is None:
        return None
    element = getattr(numbering_part, "element", None)
    if isinstance(element, etree._Element):
        return element
    return None


def _resolve_abstract_num(numbering_root: etree._Element, num_id: int) -> etree._Element | None:
    num_matches = xpath(numbering_root, f"./w:num[@w:numId='{num_id}']")
    if not num_matches:
        return None
    num_el = num_matches[0]
    abstract_ref = num_el.find(qn("w:abstractNumId"))
    if abstract_ref is None:
        return None
    abstract_id = abstract_ref.get(qn("w:val"))
    if abstract_id is None:
        return None
    abstract_matches = xpath(numbering_root, f"./w:abstractNum[@w:abstractNumId='{abstract_id}']")
    return abstract_matches[0] if abstract_matches else None


def _find_level(abstract_num: etree._Element, ilvl: int) -> etree._Element | None:
    matches = xpath(abstract_num, f"./w:lvl[@w:ilvl='{ilvl}']")
    return matches[0] if matches else None


__all__ = [
    "FormattingSource",
    "MissingPartError",
    "ResolvedFormatting",
    "StyleCascadeError",
    "resolve_effective_formatting",
]
