# Layout — columns, section breaks, line numbering, page borders

Module: `docx_plus.layout`. Page-layout operations that python-docx doesn't
expose. python-docx already handles page size, margins, orientation, and
headers/footers via `doc.sections[i]` — use it for those. `docx_plus.layout`
adds the rest.

Most functions take a `Section` (`doc.sections[0]`, or the one returned by
`insert_section_break`). The even/odd-headers toggles take the whole `doc`
because they live in `settings.xml`.

## Columns

```python
from docx_plus.layout import set_columns

set_columns(doc.sections[0], 2, space=720, separator=True)
# unequal columns (widths in twips, must match the section content width):
set_columns(doc.sections[0], 2, widths=[3000, 6000])
```

`set_columns(section, num, *, space=720, separator=False, widths=None)` — emits
`<w:cols>` into the section's `sectPr`. Idempotent (replaces any existing).
`space` is the inter-column gap in twips; `separator=True` draws a line between
columns; `widths` (a list of twip widths) gives unequal columns.

## Mid-document section breaks

python-docx can only append a section at the end. `insert_section_break` splits
at any paragraph, returning a `Section` proxy for the *new* (trailing) section —
so you can then set columns / line numbering / borders on just that part.

```python
from docx_plus.layout import insert_section_break, set_columns

doc.add_heading("Intro (single column)", level=1)
split = doc.add_paragraph("Break here")

new_section = insert_section_break(split, start_type="continuous")
set_columns(new_section, 2, separator=True)     # two-column from here on

doc.add_heading("Body (two columns)", level=1)
```

`insert_section_break(paragraph, *, start_type="nextPage")` — `start_type` is
`"nextPage"`, `"continuous"`, `"evenPage"`, `"oddPage"`, or `"nextColumn"`.

## Distinct even/odd headers

A document-level flag (distinct from per-section `titlePg`, which python-docx
already exposes as `section.different_first_page_header_footer`). After enabling
it, set `section.even_page_header` / `.odd_page_*` through python-docx.

```python
from docx_plus.layout import (
    enable_distinct_even_odd_headers, disable_distinct_even_odd_headers,
)

enable_distinct_even_odd_headers(doc)    # idempotent
# disable_distinct_even_odd_headers(doc)
```

## Line numbering

Marginal line numbers (the legal-document style).

```python
from docx_plus.layout import set_line_numbering

set_line_numbering(doc.sections[0], count_by=5, restart="newPage")
```

`set_line_numbering(section, *, count_by=1, restart="newPage", start=1, distance=None)`

- `count_by` — show every Nth number (5 ⇒ 5, 10, 15, …).
- `restart` — `"newPage"`, `"newSection"`, or `"continuous"`.
- `start` — first number.
- `distance` — gap between the number and the text, in twips (default: auto).

Idempotent; replaces any existing `<w:lnNumType>`.

## Page borders

```python
from docx_plus.layout import Border, set_page_borders

rule = Border(style="single", size=8, color="2F5496", space=24)
set_page_borders(doc.sections[0], top=rule, bottom=rule, left=rule, right=rule)
# All-None removes the page border:
# set_page_borders(doc.sections[0])
```

`set_page_borders(section, *, top=None, bottom=None, left=None, right=None)` —
one `Border` per side; omit a side to leave it borderless. Passing all `None`
removes the `<w:pgBorders>` element. Idempotent.

`Border(style, size, color, space)` (frozen dataclass):

- `style` — line style: `"single"`, `"double"`, `"dotted"`, `"dashed"`,
  `"thick"`, … (ECMA-376 `ST_Border`).
- `size` — thickness in **eighths of a point** (`8` ⇒ 1 pt).
- `color` — `"RRGGBB"` hex, or `"auto"`.
- `space` — offset from the text/page edge, in **twips**.

## End-to-end

```python
from docx import Document
from docx_plus.layout import (
    Border, enable_distinct_even_odd_headers, insert_section_break,
    set_columns, set_line_numbering, set_page_borders,
)

doc = Document()
doc.add_heading("Intro", level=1)
split = doc.add_paragraph("Two-column body starts here")

sec = insert_section_break(split, start_type="continuous")
set_columns(sec, 2, space=720, separator=True)
set_line_numbering(sec, count_by=5, restart="newPage")

rule = Border(style="single", size=8, color="2F5496", space=24)
set_page_borders(doc.sections[0], top=rule, bottom=rule, left=rule, right=rule)

enable_distinct_even_odd_headers(doc)
doc.save("layout.docx")
```

Type aliases: `SectionStartType`, `LineNumberRestart`.

See also: `docx_plus/examples/multi_column_layout.py`.
