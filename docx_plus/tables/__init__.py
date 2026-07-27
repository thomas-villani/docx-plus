"""Table formatting beyond python-docx — borders, shading, and merging.

python-docx models table *structure* well: rows, columns, cells, cell
text, widths, and a working ``_Cell.merge``. It models table
*appearance* not at all. There is no ``CT_Border``, no ``CT_TblBorders``,
no ``CT_TcBorders``, and no ``CT_Shd`` class in the package, and none of
those tags is registered — so ruling a table or shading a header row
means writing OOXML by hand.

This module fills that gap and the three merge-related ones around it:

- :func:`set_table_borders` / :func:`set_cell_borders` — ``<w:tblBorders>``
  and ``<w:tcBorders>`` over the shared
  :class:`~docx_plus.core.borders.Border` dataclass, including the two
  inside edges and the two cell diagonals.
- :func:`set_table_shading` / :func:`set_row_shading` /
  :func:`set_cell_shading` — ``<w:shd>``. Rows have no shading element
  of their own; :func:`set_row_shading` writes through to the cells, as
  Word does.
- :func:`merge_cells` — python-docx's own merge behind a typed error.
- :func:`unmerge_cell` — the inverse, which python-docx lacks entirely.
- :func:`normalize_horizontal_merges` — rewrites legacy ``<w:hMerge>``
  spans, which python-docx's grid model does not understand, into the
  ``<w:gridSpan>`` form it does.
- :func:`read_table_formatting` — reads the above back. **Direct
  formatting only**; the table-style cascade is not resolved.

This is distinct from ``layout/borders.py``, which does *page* borders
and takes a :class:`~docx.section.Section`.
"""

from __future__ import annotations

from docx_plus.tables.borders import set_cell_borders, set_table_borders
from docx_plus.tables.merge import (
    InvalidMergeError,
    merge_cells,
    normalize_horizontal_merges,
    unmerge_cell,
)
from docx_plus.tables.read import (
    CellFormatting,
    TableFormatting,
    read_table_formatting,
)
from docx_plus.tables.shading import (
    Shading,
    set_cell_shading,
    set_row_shading,
    set_table_shading,
    shading_attrs,
)

__all__ = [
    "CellFormatting",
    "InvalidMergeError",
    "Shading",
    "TableFormatting",
    "merge_cells",
    "normalize_horizontal_merges",
    "read_table_formatting",
    "set_cell_borders",
    "set_cell_shading",
    "set_row_shading",
    "set_table_borders",
    "set_table_shading",
    "shading_attrs",
    "unmerge_cell",
]
