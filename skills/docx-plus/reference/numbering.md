# Numbering — custom bullet and numbered list definitions

Module: `docx_plus.numbering`. Covers what python-docx cannot do at all: it
ships a `NumberingPart` but has no `CT_AbstractNum` and no `CT_Lvl`, so it
cannot express a number format, level text, start value, indent, or bullet
glyph.

> **The one thing to internalise:** OOXML splits a list in two. A
> `<w:abstractNum>` is the *definition*; a `<w:num>` is an *instance* of it.
> Paragraphs reference the **instance** (`numId`), never the definition. Two
> instances of one definition are independent counters that look identical —
> which is the only way to restart numbering.

## Quick start

```python
from docx import Document
from docx_plus.numbering import apply_list, define_bullet_list, define_numbered_list

doc = Document()

steps = define_numbered_list(doc, levels=2)     # 1. then a.
for text in ("First", "Second", "Third"):
    apply_list(doc.add_paragraph(text), steps)
apply_list(doc.add_paragraph("Sub-point"), steps, level=1)

bullets = define_bullet_list(doc, levels=3)     # Word's round / o / square cycle
apply_list(doc.add_paragraph("A note"), bullets)
```

Both presets return a `numId`. Pass it to `apply_list` for every paragraph
that belongs to that list; paragraphs sharing a `numId` continue one
sequence in document order.

## Custom definitions

```python
from docx_plus.numbering import LevelDefinition, define_list_definition

# A legal outline: 1. / 1.1. / 1.1.1.
outline = define_list_definition(
    doc,
    levels=[
        LevelDefinition(text="%1.",         indent=720,  hanging=360),
        LevelDefinition(text="%1.%2.",      indent=1584, hanging=504),
        LevelDefinition(text="%1.%2.%3.",   indent=2448, hanging=792),
    ],
    name="Sign-off outline",
)
```

`%N` interpolates the counter for level `N`, **1-based** — so level 0's own
counter is `%1`. A level can only reference counters at or above its own
depth; `%3` on level 0 raises `InvalidLevelError`.

`LevelDefinition` fields: `fmt` (any ECMA-376 `ST_NumberFormat` — `decimal`,
`lowerLetter`, `lowerRoman`, `upperRoman`, `bullet`, …), `text`, `start`,
`indent`, `hanging`, `justify`, `suffix` (`tab` / `space` / `nothing`),
`restart_after`, `font`. All validated at construction.

### Two traps

**Size `hanging` to the number.** It is the width reserved for the number,
and the gap to the text is a tab stop at `indent`. If the number is wider
than `hanging` the tab collapses and you get `1.1.1.On-call lead` instead of
`1.1.1. On-call lead`. Deeper outline levels need progressively larger
values — this is visible only in Word, never in the XML.

**Symbol bullets need a font.** `LevelDefinition(fmt="bullet", text="")`
renders as a Latin letter unless you also pass `font="Symbol"`. The presets
handle this for you.

## Restarting

```python
from docx_plus.numbering import restart_list

second = restart_list(doc.add_paragraph("One again"), steps, start=1)
apply_list(doc.add_paragraph("Two again"), second)
```

`restart_list` returns a **new** `numId`. Apply it to every paragraph in the
restarted run — the call only moves the paragraph you pass.

## Removing

```python
from docx_plus.numbering import remove_list

remove_list(p)                                    # drop direct numbering
remove_list(p, suppress_style_numbering=True)     # also defeat the style's
```

Plain `remove_list` reverts the paragraph to whatever its *style* says — and
for `List Bullet` that means it stays bulleted. The flag writes
`<w:numPr><w:numId w:val="0"/></w:numPr>`, Word's sentinel for "definitely
not numbered".

## Reading back

```python
from docx_plus.numbering import read_list_definitions

for d in read_list_definitions(doc):
    print(d.num_id, d.name, [lvl.fmt for lvl in d.levels], d.start_overrides)
```

Returns `[]` when the document has no `numbering.xml`, and never creates
one. `start_overrides` is what distinguishes a restarted instance from the
original.

**A fresh `Document()` is not empty.** python-docx's bundled template ships
nine definitions backing `List Bullet` / `List Number`, so an untouched
document reports nine. Filter by the `numId` you were given.

## Sharing allocators

```python
from docx_plus.numbering import AbstractNumIdRegistry, NumIdRegistry

nums, abstracts = NumIdRegistry(doc), AbstractNumIdRegistry(doc)
a = define_numbered_list(doc, num_registry=nums, abstract_registry=abstracts)
b = define_bullet_list(doc, num_registry=nums, abstract_registry=abstracts)
```

Optional but worth it when defining several lists in one pass: the registries
allocate the lowest free id and seed from what is already in the file, so ids
stay dense and never collide with the template's.

## Gotchas

- `apply_list` does **not** validate `num_id`. A dangling reference is legal
  OOXML and Word renders the paragraph unnumbered.
- Nine levels maximum per definition (`MAX_LEVELS`).
- `abstractNumId` starts at 0; `numId` starts at 1, because `numId` 0 is the
  "no numbering" sentinel.
- A `ListBullet` style created by `docx_plus.styles.ensure_style` carries no
  numbering of its own. Apply a definition to the paragraphs directly.
