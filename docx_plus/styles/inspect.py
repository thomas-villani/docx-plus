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
from docx_plus.styles.theme import (
    ThemeColors,
    load_theme,
    resolve_theme_color,
    resolve_theme_font,
)

if TYPE_CHECKING:
    from docx.document import Document
    from docx.table import _Cell
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run


_MAX_STYLE_CHAIN_DEPTH = 11


@dataclass(frozen=True)
class TableContext:
    """A cell's position within its table — for conditional table-style formatting.

    ECMA-376 17.7.6.5 lets a ``<w:style w:type="table">`` carry
    conditional formatting branches (``<w:tblStylePr w:type="firstRow"/>``,
    ``"lastRow"``, ``"firstCol"``, ``"lastCol"``, ``"band1Horz"``,
    ``"band1Vert"``, ``"band2Horz"``, ``"band2Vert"``,
    ``"nwCell"`` / ``"neCell"`` / ``"swCell"`` / ``"seCell"``). To pick
    the right branches the cascade resolver needs to know where in the
    table the target lives.

    Construct manually for an out-of-band query, or pass a ``_Cell`` to
    :func:`resolve_effective_formatting` to derive the context
    automatically from the cell's parent row / table.

    Band size: by default rows alternate band1 / band2 every row. When
    the table instance's ``<w:tblPr>`` carries a ``<w:tblStyleRowBandSize
    w:val="N"/>`` (resp. ``<w:tblStyleColBandSize>``), bands span ``N``
    rows / columns each. Note that v0.2 does not yet walk the table
    **style chain** looking for these attributes — only the table
    instance's own ``tblPr`` is consulted. This is sufficient for tables
    where the application or user explicitly set the band size, but
    misses style-defined band sizes (deferred to v0.3+).

    Scope: this context selects which ``<w:tblStylePr>`` branches apply,
    but only their **run / paragraph** properties are resolved. Cell-,
    row-, and table-level properties (``<w:tcPr>`` / ``<w:trPr>`` /
    ``<w:tblPr>``) from a table style are not surfaced — see the
    :func:`resolve_effective_formatting` note.

    Auto-derivation limitation: when a row wraps its cells in a
    ``<w:sdt>`` (a content control around table cells), the derived
    column index cannot be computed and an empty (all-False)
    :class:`TableContext` is returned. Pass an explicit context in that
    case. Nested tables resolve against the **inner** cell's position.

    Attributes:
        is_first_row: Cell is in the first ``<w:tr>`` of its table.
        is_last_row: Cell is in the last ``<w:tr>``.
        is_first_col: Cell is the first ``<w:tc>`` of its row.
        is_last_col: Cell is the last ``<w:tc>`` of its row.
        is_band_row: Cell is in a "band1" horizontal stripe (first band).
        is_band_col: Cell is in a "band1" vertical stripe (first band).
        is_band2_row: Cell is in a "band2" horizontal stripe (second
            band — the complement of band1 at default band-size=1).
        is_band2_col: Cell is in a "band2" vertical stripe.
    """

    is_first_row: bool = False
    is_last_row: bool = False
    is_first_col: bool = False
    is_last_col: bool = False
    is_band_row: bool = False
    is_band_col: bool = False
    is_band2_row: bool = False
    is_band2_col: bool = False


# ``<w:tblStylePr w:type=...>`` values in ECMA-376 17.7.6.5 application
# order: later entries override earlier ones. ``wholeTable`` always
# applies; the rest depend on the resolver's :class:`TableContext`.
# Rows precede columns so that column branches win at row/col intersections
# (which is why corner branches exist as the final override layer).
_TBL_STYLE_PR_ORDER: tuple[str, ...] = (
    "wholeTable",
    "band1Vert",
    "band2Vert",
    "band1Horz",
    "band2Horz",
    "firstRow",
    "lastRow",
    "firstCol",
    "lastCol",
    "nwCell",
    "neCell",
    "swCell",
    "seCell",
)

# Toggle properties XOR through the cascade per ECMA-376 17.7.3.
# Mapped from rPr child name to ResolvedFormatting field name.
_TOGGLE_RPR: dict[str, str] = {
    "b": "bold",
    "i": "italic",
    "bCs": "cs_bold",
    "iCs": "cs_italic",
    "caps": "caps",
    "smallCaps": "small_caps",
    "strike": "strike",
    "vanish": "vanish",
    "emboss": "emboss",
    "imprint": "imprint",
    "outline": "outline",
    "shadow": "shadow",
}
# dstrike is intentionally excluded from the XOR toggle set — per
# ECMA-376 17.7.3 the toggle list is the twelve above. dstrike is handled
# in :func:`_apply_rpr` as a non-toggle property (last writer wins) and
# surfaced on :class:`ResolvedFormatting.double_strike`.


Layer = Literal[
    "docDefaults",
    "tableStyle",
    "paragraphStyle",
    "linkedCharStyle",
    "numbering",
    "directParagraph",
    "runStyle",
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

    All twelve ECMA-376 17.7.3 toggle properties are surfaced: the six
    base toggles (``bold``, ``italic``, ``caps``, ``small_caps``,
    ``strike``, ``vanish``) and the six complex-script / decorative
    variants (``cs_bold``, ``cs_italic``, ``emboss``, ``imprint``,
    ``outline``, ``shadow``). All XOR through the cascade with the same
    semantics; an explicit ``w:val="false"`` resets parity to false.
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
    cs_bold: bool | None = None
    cs_italic: bool | None = None
    underline: str | None = None
    strike: bool | None = None
    double_strike: bool | None = None
    color_rgb: str | None = None
    highlight: str | None = None
    caps: bool | None = None
    small_caps: bool | None = None
    vanish: bool | None = None
    emboss: bool | None = None
    imprint: bool | None = None
    outline: bool | None = None
    shadow: bool | None = None
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
    table_context: TableContext | None = None,
) -> ResolvedFormatting:
    """Resolve the effective formatting for ``target``.

    Walks the six cascade layers in precedence order, returning a fully
    resolved :class:`ResolvedFormatting`. Toggle properties XOR through the
    chain per ECMA-376 17.7.3. Theme colors are resolved against the
    document's theme part; if the theme is missing or malformed, the result's
    ``partial`` flag is set and unresolved theme names are returned in place
    of hex values.

    When ``target`` is in a table cell, table-style **conditional
    formatting** (``<w:tblStylePr>`` branches: ``firstRow``, ``lastRow``,
    ``firstCol``, ``lastCol``, ``band1Horz``, ``band1Vert``, the four
    corners, and ``wholeTable``) is applied on top of the base table
    style in ECMA-376 17.7.6.5 precedence order.

    Note:
        Only **run- and paragraph-level** properties are resolved (the
        ``<w:rPr>`` / ``<w:pPr>`` carried by a style's base and its
        ``<w:tblStylePr>`` branches). Cell-, row-, and table-level
        properties (``<w:tcPr>`` cell shading and margins, ``<w:trPr>``
        row heights, ``<w:tblPr>`` table defaults) declared by a table
        style are **not** surfaced on :class:`ResolvedFormatting` — that
        belongs to a separate cell-formatting resolver deferred to v0.3+.

    Args:
        target: A python-docx :class:`~docx.text.paragraph.Paragraph`,
            :class:`~docx.text.run.Run`, or :class:`~docx.table._Cell`.
        include_provenance: If True, populate ``.provenance`` with the cascade
            layer that set each field. Default False.
        table_context: Optional override for the cell's position within
            its table. When ``None`` (default), the resolver derives it
            from the target's parent ``<w:tr>`` / ``<w:tbl>`` chain;
            pass an explicit :class:`TableContext` to query a hypothetical
            position (e.g. "what would the formatting be if this cell
            were in the first row?").

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
    target_kind, target_el = _classify_target(target)
    doc = _document_of(target)
    styles_root = doc.styles.element
    theme = load_theme(doc)

    # ``partial`` is set lazily — only when a theme reference actually fails
    # to resolve (inside _resolve_color / _resolve_font_theme). A missing
    # theme part is not, on its own, an incomplete resolution: a document
    # with no theme refs resolves fully even without a theme (SPEC §4).
    acc = _Accumulator(theme=theme, want_provenance=include_provenance)

    # _classify_target returns the underlying element alongside the kind, so
    # the union-attr access happens once where isinstance has already narrowed
    # the type — no per-branch type: ignore needed here.
    if target_kind == "paragraph":
        ctx = table_context or _derive_table_context_from_element(target_el)
        _apply_paragraph_cascade(acc, doc, styles_root, target_el, table_context=ctx)
    elif target_kind == "run":
        paragraph_element = _enclosing_paragraph(target_el)
        ctx = table_context or _derive_table_context_from_element(paragraph_element)
        _apply_paragraph_cascade(
            acc,
            doc,
            styles_root,
            paragraph_element,
            run_element=target_el,
            table_context=ctx,
        )
    else:  # cell
        ctx = table_context or _derive_table_context_from_element(target_el)
        _apply_cell_cascade(acc, styles_root, target_el, table_context=ctx)

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
    table_context: TableContext | None = None,
) -> None:
    """Walk layers 1, 3, 4, 5 (and 6 if run_element) for a paragraph target."""
    # Layer 1: docDefaults
    _apply_doc_defaults(acc, styles_root)

    # Layer 2: table style (if inside a table)
    enclosing_tc = _enclosing_cell(p_element)
    if enclosing_tc is not None:
        table_element = _enclosing_table(enclosing_tc)
        if table_element is not None:
            _apply_table_style_chain(acc, styles_root, table_element, table_context=table_context)

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

        # Run-level rStyle reference (character style applied to one run).
        # Per ECMA-376 17.3.2.29 this is a style layer that sits BELOW direct
        # run formatting — direct rPr on the run must override it.
        run_style_id = _run_style_id(run_element)
        if run_style_id is not None:
            _apply_style_chain(acc, styles_root, run_style_id, "runStyle")

        # Layer 6: direct run formatting (highest precedence for the run).
        run_rpr = run_element.find(qn("w:rPr"))
        if run_rpr is not None:
            _apply_rpr(acc, run_rpr, FormattingSource(layer="directRun"))


def _apply_cell_cascade(
    acc: _Accumulator,
    styles_root: etree._Element,
    tc_element: etree._Element,
    table_context: TableContext | None = None,
) -> None:
    """Resolve formatting for a table cell — table style chain only, for now.

    Takes no ``doc`` parameter (unlike :func:`_apply_paragraph_cascade`):
    cells carry no paragraph-level numbering, so the doc-aware numbering
    layer that needs ``numbering.xml`` does not apply here.
    """
    _apply_doc_defaults(acc, styles_root)
    table_element = _enclosing_table(tc_element)
    if table_element is not None:
        _apply_table_style_chain(acc, styles_root, table_element, table_context=table_context)


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
            # basedOn is single-valued, so the chain is linear: this path is
            # the real basedOn sequence up to the repeat, never a diamond.
            # A self-cycle (X basedOn X) prints as "X -> X".
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
    table_context: TableContext | None = None,
) -> None:
    """Apply the table's style chain in spec-correct interleaved order.

    Walks the basedOn chain ONCE, ancestors-first. For each style level
    apply its base ``pPr`` / ``rPr`` then — when a
    :class:`TableContext` is provided — its matching
    ``<w:tblStylePr w:type="...">`` branches in ECMA-376 17.7.6.5
    precedence order (``wholeTable`` → bands → first/last row →
    first/last col → corners). This ensures the per-level invariant
    "conditional branches override that level's base" holds while
    still letting a child level's everything (base + conditional)
    override a parent level's everything.
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

    chain = _collect_style_chain(styles_root, style_id)
    matching = _matching_conditional_types(table_context) if table_context is not None else set()

    # Ancestors-first: reverse the leaf-to-root chain.
    for depth, (sid, style_el) in enumerate(reversed(chain)):
        chain_depth = len(chain) - 1 - depth
        source = FormattingSource(layer="tableStyle", style_id=sid, chain_depth=chain_depth)

        # 1. Base pPr / rPr for this style level.
        ppr = style_el.find(qn("w:pPr"))
        if ppr is not None:
            _apply_ppr(acc, ppr, source)
        rpr = style_el.find(qn("w:rPr"))
        if rpr is not None:
            _apply_rpr(acc, rpr, source)

        # 2. Matching conditional branches for this style level, in spec order.
        if not matching:
            continue
        branches: dict[str, etree._Element] = {}
        for branch in style_el.findall(qn("w:tblStylePr")):
            type_attr = branch.get(qn("w:type"))
            if type_attr is not None:
                branches[type_attr] = branch
        for cond_type in _TBL_STYLE_PR_ORDER:
            if cond_type not in matching or cond_type not in branches:
                continue
            branch = branches[cond_type]
            branch_ppr = branch.find(qn("w:pPr"))
            if branch_ppr is not None:
                _apply_ppr(acc, branch_ppr, source)
            branch_rpr = branch.find(qn("w:rPr"))
            if branch_rpr is not None:
                _apply_rpr(acc, branch_rpr, source)


def _matching_conditional_types(ctx: TableContext) -> set[str]:
    """Return the set of ``<w:tblStylePr w:type=...>`` values that apply.

    ``wholeTable`` always matches. Each positional flag activates its
    corresponding type, and corner types match only when both axes
    align.
    """
    types: set[str] = {"wholeTable"}
    if ctx.is_band_col:
        types.add("band1Vert")
    if ctx.is_band2_col:
        types.add("band2Vert")
    if ctx.is_band_row:
        types.add("band1Horz")
    if ctx.is_band2_row:
        types.add("band2Horz")
    if ctx.is_first_col:
        types.add("firstCol")
    if ctx.is_last_col:
        types.add("lastCol")
    if ctx.is_first_row:
        types.add("firstRow")
    if ctx.is_last_row:
        types.add("lastRow")
    if ctx.is_first_row and ctx.is_first_col:
        types.add("nwCell")
    if ctx.is_first_row and ctx.is_last_col:
        types.add("neCell")
    if ctx.is_last_row and ctx.is_first_col:
        types.add("swCell")
    if ctx.is_last_row and ctx.is_last_col:
        types.add("seCell")
    return types


def _read_band_size(tbl: etree._Element, child_name: str) -> int:
    """Read ``<w:tblStyleRowBandSize>`` / ``<w:tblStyleColBandSize>`` from a table.

    Looks at the table instance's own ``<w:tblPr>``. Returns 1 if the
    element is absent or the value is unparseable. Does not currently
    walk the table style chain — see :class:`TableContext` docstring.
    """
    tbl_pr = tbl.find(qn("w:tblPr"))
    if tbl_pr is None:
        return 1
    el = tbl_pr.find(qn(child_name))
    if el is None:
        return 1
    raw = el.get(qn("w:val"))
    if raw is None:
        return 1
    try:
        n = int(raw)
    except ValueError:
        return 1
    return n if n >= 1 else 1


def _derive_table_context_from_element(node: etree._Element) -> TableContext:
    """Derive a :class:`TableContext` from a body element's table position.

    Walks up from ``node`` to find the enclosing ``<w:tc>``, then derives
    row / column indices and band parity. Returns an empty (all-False)
    :class:`TableContext` when ``node`` is not inside a table.

    Band parity follows the convention that row index 1, 3, 5, ... is
    "band1" and 2, 4, 6, ... is "band2" at the default band-size of 1.
    Row 0 is treated as not-banded; in practice the ``firstRow``
    conditional (if defined) overrides any band branch at row 0 per
    ECMA-376 17.7.6.5 precedence. When the table's ``<w:tblPr>``
    declares ``tblStyleRowBandSize`` or ``tblStyleColBandSize``, bands
    span that many rows / columns each.
    """
    if isinstance(node.tag, str) and etree.QName(node.tag).localname == "tc":
        tc: etree._Element | None = node
    else:
        tc = _enclosing_cell(node)
    if tc is None:
        return TableContext()
    tr = tc.getparent()
    if tr is None or tr.tag != qn("w:tr"):
        return TableContext()
    tbl = tr.getparent()
    if tbl is None or tbl.tag != qn("w:tbl"):
        return TableContext()

    rows = [child for child in tbl if child.tag == qn("w:tr")]
    cells = [child for child in tr if child.tag == qn("w:tc")]
    try:
        row_idx = rows.index(tr)
        col_idx = cells.index(tc)
    except ValueError:
        # tr/tc not a direct child of its parent — happens when a <w:sdt>
        # wraps the row's cells. Position is indeterminate; fall back to an
        # empty context (caller may pass an explicit one). See TableContext.
        return TableContext()

    row_band_size = _read_band_size(tbl, "w:tblStyleRowBandSize")
    col_band_size = _read_band_size(tbl, "w:tblStyleColBandSize")

    # Row 0 is excluded from the banding sequence (firstRow's job, if it
    # exists). For row_idx >= 1, the (row_idx - 1) // size yields the
    # stripe index; even stripes are band1, odd are band2.
    if row_idx >= 1:
        row_stripe = (row_idx - 1) // row_band_size
        is_band_row = (row_stripe % 2) == 0
        is_band2_row = (row_stripe % 2) == 1
    else:
        is_band_row = False
        is_band2_row = False

    if col_idx >= 1:
        col_stripe = (col_idx - 1) // col_band_size
        is_band_col = (col_stripe % 2) == 0
        is_band2_col = (col_stripe % 2) == 1
    else:
        is_band_col = False
        is_band2_col = False

    return TableContext(
        is_first_row=row_idx == 0,
        is_last_row=row_idx == len(rows) - 1,
        is_first_col=col_idx == 0,
        is_last_col=col_idx == len(cells) - 1,
        is_band_row=is_band_row,
        is_band_col=is_band_col,
        is_band2_row=is_band2_row,
        is_band2_col=is_band2_col,
    )


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
            # A theme token (w:asciiTheme) resolves against the theme's
            # font scheme; a literal face (w:ascii / w:hAnsi / w:cs) is used
            # verbatim. Theme attributes take precedence — that is what Word
            # writes when a font is theme-bound.
            ascii_theme = child.get(qn("w:asciiTheme"))
            if ascii_theme is not None:
                acc.set("font_name", _resolve_font_theme(ascii_theme, acc), source)
            else:
                literal = (
                    child.get(qn("w:ascii")) or child.get(qn("w:hAnsi")) or child.get(qn("w:cs"))
                )
                if literal is not None:
                    acc.set("font_name", literal, source)
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
        elif local == "dstrike":
            # ECMA-376 17.3.2.10: not a toggle (last writer wins).
            val = child.get(qn("w:val"))
            acc.set("double_strike", val != "false" and val != "0", source)
        elif local == "highlight":
            val = child.get(qn("w:val"))
            if val is not None:
                acc.set("highlight", val, source)
        elif local == "vertAlign":
            val = child.get(qn("w:val"))
            if val is not None:
                acc.set("vert_align", val, source)


def _resolve_color(color_el: etree._Element, acc: _Accumulator) -> str | None:
    """Resolve a ``<w:color>`` element to an uppercase ``RRGGBB`` hex string.

    Handles the two theme transforms ``<w:color>`` can carry —
    ``themeTint`` and ``themeShade`` (ECMA-376 CT_Color). The DrawingML
    ``lumMod`` / ``lumOff`` transforms are not applicable here: the
    ``w:color`` schema cannot carry them (see :mod:`docx_plus.styles.theme`).

    On an unresolvable theme reference the result is flagged
    ``partial`` (SPEC §4). The unresolved name is surfaced as the value
    **only** when the theme part is entirely absent — so callers without a
    theme can still log which color was wanted. When the theme loaded but
    the name is not in its scheme (a typo such as ``"accent7"``, or the
    explicit ``"none"`` sentinel), no value is returned: a bare name would
    land a non-hex string in ``color_rgb`` that the style writers reject.
    """
    theme_name = color_el.get(qn("w:themeColor"))
    if theme_name is not None:
        if theme_name == "none":
            # Explicit "no theme color" — not a resolution failure.
            return None
        tint = color_el.get(qn("w:themeTint"))
        shade = color_el.get(qn("w:themeShade"))
        resolved = resolve_theme_color(acc.theme, theme_name, tint=tint, shade=shade)
        if resolved is not None:
            return resolved
        acc.partial = True
        if acc.theme is None:
            return theme_name
        return None
    val = color_el.get(qn("w:val"))
    if val and val.lower() != "auto":
        return val.upper()
    return None


def _resolve_font_theme(token: str, acc: _Accumulator) -> str:
    """Resolve a ``w:asciiTheme`` font token to its concrete typeface.

    Reads the theme's ``a:fontScheme`` (e.g. ``"minorHAnsi"`` -> ``"Calibri"``).
    When the theme is absent or the token has no scheme entry the token is
    surfaced unchanged and the result is flagged ``partial`` — the same
    contract :func:`_resolve_color` uses, so a ``partial=True`` result
    reliably means "a theme reference did not resolve to a concrete value"
    (SPEC §4).
    """
    resolved = resolve_theme_font(acc.theme, token)
    if resolved is not None:
        return resolved
    acc.partial = True
    return token


# --------------------------------------------------------------------------
# Document traversal / metadata helpers.
# --------------------------------------------------------------------------


def _classify_target(
    target: object,
) -> tuple[Literal["paragraph", "run", "cell"], etree._Element]:
    """Classify ``target`` and return ``(kind, underlying_element)``.

    Returning the element here — where ``isinstance`` has narrowed the type
    — lets the caller avoid a ``type: ignore[union-attr]`` on each of
    ``._p`` / ``._r`` / ``._tc``.
    """
    from docx.table import _Cell
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run

    if isinstance(target, Paragraph):
        return "paragraph", target._p
    if isinstance(target, Run):
        return "run", target._r
    if isinstance(target, _Cell):
        return "cell", target._tc
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
    matches = xpath(styles_root, "./w:style[@w:styleId=$sid]", sid=style_id)
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
    raise StyleCascadeError("run is not inside a paragraph")


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
    num_matches = xpath(numbering_root, "./w:num[@w:numId=$nid]", nid=str(num_id))
    if not num_matches:
        return None
    num_el = num_matches[0]
    abstract_ref = num_el.find(qn("w:abstractNumId"))
    if abstract_ref is None:
        return None
    abstract_id = abstract_ref.get(qn("w:val"))
    if abstract_id is None:
        return None
    abstract_matches = xpath(
        numbering_root,
        "./w:abstractNum[@w:abstractNumId=$aid]",
        aid=abstract_id,
    )
    return abstract_matches[0] if abstract_matches else None


def _find_level(abstract_num: etree._Element, ilvl: int) -> etree._Element | None:
    matches = xpath(abstract_num, "./w:lvl[@w:ilvl=$lvl]", lvl=str(ilvl))
    return matches[0] if matches else None


__all__ = [
    "FormattingSource",
    "MissingPartError",
    "ResolvedFormatting",
    "StyleCascadeError",
    "TableContext",
    "resolve_effective_formatting",
]
