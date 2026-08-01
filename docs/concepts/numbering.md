# Custom numbering

`numbering/` (v0.5) closes the largest remaining python-docx gap.
python-docx ships a `NumberingPart` and `len()` of its definitions, and
that is all: `docx/oxml/numbering.py` defines `CT_Numbering`, `CT_Num`,
`CT_NumLvl`, and `CT_NumPr`, but there is **no `CT_AbstractNum` and no
`CT_Lvl`**. Nothing in it can express a number format, level text, start
value, indent, or bullet glyph, so building a list means hand-writing
XML.

For the calls, see the [numbering guide](../guides/numbering.md).

## The two-part model

```
numbering.xml
├── <w:abstractNum w:abstractNumId="0">    the definition
│     └── <w:lvl w:ilvl="0..8">            one per outline depth
└── <w:num w:numId="1">                    an instance
      └── <w:abstractNumId w:val="0"/>
```

A paragraph's `<w:numPr>` names the **instance**, never the abstract
definition. That indirection is not incidental — it is the whole
mechanism behind restarting. `restart_list` adds a *second* `<w:num>`
over the same `<w:abstractNum>` carrying a `<w:startOverride>`, giving an
independent counter that renders identically. There is nowhere in OOXML
to mark a paragraph "count from 1 again"; this is how Word does it too.

## Three things that must be right

1. **`w:abstractNum` precedes every `w:num`.** `CT_Numbering` is
   `numPicBullet*, abstractNum*, num*, numIdMacAtCleanup?`, and since
   nothing in python-docx inserts an `abstractNum` at all, the ordering
   is entirely on us — `define.py` uses
   `insert_before_first_anchor(root, node, ("w:num", "w:numIdMacAtCleanup"))`.
   Getting it wrong produces a file lenient parsers accept and Word may
   not. Pinned by `assert_numbering_well_formed`, which also checks id
   uniqueness and that every instance resolves.
2. **`w:lvl` children follow ECMA-376 17.9.6 order** —
   `start, numFmt, lvlRestart, pStyle, isLgl, suff, lvlText,
   lvlPicBulletId, legacy, lvlJc, pPr, rPr`. `_LVL_CHILD_ORDER` plus the
   promoted `core.ordered_insert` handle it.
3. **Symbol bullets need their font.** `U+F0B7` and `U+F0A7` are
   private-use codepoints; without `w:rPr/w:rFonts` naming Symbol or
   Wingdings — *and* `w:hint="default"` — Word substitutes a theme font
   and renders the bullet as a Latin letter.

## Allocation

Both id namespaces use `next_sequential()` rather than the random
`next()` the rest of the library uses: Word and python-docx both take the
lowest free integer, and a `numbering.xml` full of nine-digit ids is
needlessly unreadable. `AbstractNumIdRegistry` lowers `_MIN_ID` to `0`
because `w:abstractNumId` legitimately starts there, while `NumIdRegistry`
stays at `1` — inside a `w:numPr`, `numId` `0` is the sentinel for "no
numbering", which is the only way a paragraph opts out of a list applied
by its style.

## The indent trap

`hanging` is the width reserved for the number, and the gap between
number and text is a tab stop sitting at `indent`. When the number is
wider than `hanging` the tab has nowhere to advance to and collapses, so
a cumulative outline renders `1.1.1.On-call lead`. Deeper levels of a
`%1.%2.%3.` outline need progressively larger hanging values. Found by
opening the example in Word; the suite cannot see it.

## Not covered

Linking a definition into a *style* (`w:style/w:pPr/w:numPr`).
`styles/modify.py` already owns writing into `w:style` and carries
`_STYLE_CHILD_ORDER` for it, so that belongs there rather than duplicating
the schema knowledge here. Until it lands, a `ListBullet` style created by
`ensure_style` carries no numbering of its own (`modify.py` says so at its
`_BUILTIN_STYLES` definition) — apply a definition to the paragraphs
directly. Tracked in `ROADMAP.md`.
