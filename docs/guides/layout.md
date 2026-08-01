# Page layout

`docx_plus.layout` adds the page-layout operations python-docx doesn't
expose. It already handles page size, margins, orientation, and
headers/footers through `doc.sections[i]` — keep using it for those.

Most functions here take a `Section` (`doc.sections[0]`, or the one
returned by `insert_section_break`). The even/odd-header toggles take the
whole `doc`, because that flag lives in `settings.xml`.

## Columns

```python
from docx_plus.layout import set_columns

set_columns(doc.sections[0], 2, space=720, separator=True)

# Unequal columns — widths in twips, summing to the section content width:
set_columns(doc.sections[0], 2, widths=[3000, 6000])
```

`set_columns(section, num, *, space=720, separator=False, widths=None)` —
`space` is the inter-column gap in twips, `separator=True` draws a rule
between columns. Idempotent: replaces any existing `<w:cols>`.

## Mid-document section breaks

python-docx can only append a section at the end of a document.
`insert_section_break` splits at any paragraph and returns a `Section` proxy
for the **new (trailing) section** — so you can then set columns, line
numbering, or borders on just that part.

```python
from docx_plus.layout import insert_section_break, set_columns

doc.add_heading("Intro (single column)", level=1)
split = doc.add_paragraph("Break here")

new_section = insert_section_break(split, start_type="continuous")
set_columns(new_section, 2, separator=True)     # two columns from here on

doc.add_heading("Body (two columns)", level=1)
```

`insert_section_break(paragraph, *, start_type="nextPage")` — `start_type`
is `"nextPage"`, `"continuous"`, `"evenPage"`, `"oddPage"`, or
`"nextColumn"`.

The new section inherits everything (page size, margins, header/footer
references) from the document's trailing section properties, so both halves
render identically until you change one.

## Distinct even/odd headers

```python
from docx_plus.layout import (
    disable_distinct_even_odd_headers,
    enable_distinct_even_odd_headers,
)

enable_distinct_even_odd_headers(doc)     # idempotent
```

!!! note "Three different things get confused here"
    | Setting | Scope | Where |
    |---|---|---|
    | `evenAndOddHeaders` | Document | This function |
    | `titlePg` — distinct *first*-page header | Section | python-docx's `section.different_first_page_header_footer` |
    | The `even` header/footer reference itself | Section | python-docx's `section.even_page_header` |

    A real even-page-distinct workflow needs all three: enable the
    doc-level flag here, then set `section.even_page_header` through
    python-docx.

## Line numbering

Marginal line numbers — the legal and contract document style.

```python
from docx_plus.layout import set_line_numbering

set_line_numbering(doc.sections[0], count_by=5, restart="newPage")
```

`set_line_numbering(section, *, count_by=1, restart="newPage", start=1,
distance=None)`

- `count_by` — show every Nth number (`5` gives 5, 10, 15, …)
- `restart` — `"newPage"`, `"newSection"`, or `"continuous"`
- `start` — the first number
- `distance` — gap between the number and the text, in twips (default auto)

Idempotent. `restart` validates eagerly; `count_by` and `start` must be ≥ 1.

## Page borders

```python
from docx_plus.layout import Border, set_page_borders

rule = Border(style="single", size=8, color="2F5496", space=24)
set_page_borders(doc.sections[0], top=rule, bottom=rule, left=rule, right=rule)

set_page_borders(doc.sections[0])     # all None -> removes the border
```

`set_page_borders(section, *, top=None, bottom=None, left=None,
right=None)` — one `Border` per side; omit a side to leave it borderless.
Passing all four as `None` removes the `<w:pgBorders>` element rather than
emitting an empty one. Idempotent.

`Border(style, size, color, space)` is a frozen dataclass:

| Field | Meaning |
|---|---|
| `style` | ECMA-376 `ST_Border` line style: `"single"`, `"double"`, `"dotted"`, `"dashed"`, `"thick"`, … |
| `size` | Thickness in **eighths of a point** (`8` is 1pt) |
| `color` | `"RRGGBB"` hex, or `"auto"` |
| `space` | Offset from the text or page edge, in **twips** |

The same `Border` drives [table and cell borders](tables.md) — but note
that `space` behaves differently there.

## End to end

```python
from docx import Document
from docx_plus.layout import (
    Border,
    enable_distinct_even_odd_headers,
    insert_section_break,
    set_columns,
    set_line_numbering,
    set_page_borders,
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

## See also

- [How layout works](../concepts/layout.md)
- Reference: [`layout.columns`](../reference/layout-columns.md),
  [`layout.breaks`](../reference/layout-breaks.md),
  [`layout.settings`](../reference/layout-settings.md),
  [`layout.line_numbering`](../reference/layout-line-numbering.md),
  [`layout.borders`](../reference/layout-borders.md)
- Example: `multi_column_layout.py`
