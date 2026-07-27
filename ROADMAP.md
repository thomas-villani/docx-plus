# docx_plus — Roadmap

The single authoritative roadmap for `docx_plus`. `SPEC.md §15` (a v0.1-era
historical list) and `docs/ARCHITECTURE.md §11` both defer to this file.

`docx_plus` is, and stays, a lean extension to `python-docx` that does the
things `python-docx` can't. Every item below either fills a documented
`python-docx` gap or rounds out a surface already started here. Ideas that
don't fit that charter are routed to sibling projects, not absorbed.

## Current state — v0.4.0 released

Tagged: `v0.1.0`, `v0.2.0`, `v0.2.1`, `v0.3.0` (2026-06-15), `v0.4.0`
(2026-07-26). Shipped capability modules:

| Module | Surface |
|---|---|
| `styles/` | Cascade inspection (`ResolvedFormatting`, all 12 toggles, theme fonts + colors, conditional table styles), modification, remapping |
| `controls/` | Content controls — `FormBuilder`, read / set / clear values |
| `fields/` | Simple + complex fields, `mark_fields_dirty` |
| `protection/` | `protect_document` |
| `comments/` | Anchored comments — add / edit / delete / clear, over runs, paragraphs, run ranges; threads — reply, resolve / reopen, nested read (v0.4); durable ids + author presence (v0.5) |
| `layout/` | Columns, mid-document section breaks, even/odd headers, line numbering, page borders |
| `bookmarks/` | Bookmarks + `REF` / `PAGEREF` cross-references |
| `notes/` | Footnotes + endnotes — add / edit / read |
| `publishing/` | TOC, captions, table of figures |
| `tables/` | Table / cell borders, table / row / cell shading, merge + unmerge, `w:hMerge` normalization, direct-formatting read (v0.5) |
| `revisions/` | Tracked changes — mark insertions / deletions, read revisions, accept / reject, track-changes toggle (v0.3) |
| `cli/` | `docx-plus` console command — `inspect` (effective formatting), `restyle` (style remapping), `controls` (list / set / clear values) (v0.3), `comments` (list / resolve / reopen threads) (v0.4) |

Suite at the v0.4.0 release: 905 tests (895 pass, 10 LibreOffice-skipped),
94% coverage; `mypy --strict`, `ruff`, and `mkdocs build --strict` all
clean.

## v0.5 — in progress

### Comment durable ids + author presence — shipped

The two side-parts split out of the v0.4 cycle. Both were verified
against a file **Word 2016 authored itself** — driven over COM, saved,
unzipped, and read — before a line was written, which was the right call
twice over:

- **All four content-type / relationship URIs guessed in the groundwork
  PR turned out exact.** No corrections needed.
- **`w16cid:durableId` is hex, not decimal.** The plan specified a
  decimal collector and a decimal registry. Word writes
  `ST_LongHexNumber` — the same 8-uppercase-digit form as
  `w14:paraId` — so `DurableIdRegistry` reuses the existing hex
  machinery and the feature added *less* core code than budgeted.
- **`w15:person` is not a bare element.** It carries a
  `<w15:presenceInfo>` child with `providerId` and `userId`.

Durable ids are written automatically, because they are a comment's only
identifier stable across edits and Word regenerates missing ones anyway.
Author presence is **opt-in** (`set_author_presence`): it is cosmetic,
and registering an author means inventing a `userId` for someone the
library knows nothing about.

Round-trip verified: Word opened a document written here and resaved it
preserving both `paraId` and `durableId` byte-for-byte, plus a
non-default `providerId="AD"` presence entry, and added no parts of its
own.

Deferred: `commentsExtensible.xml` — see the backlog.

### Table borders / shading / merging — shipped

Landed as a new `tables/` capability. The backlog entry bundled three
things at very different states, which is worth recording:

- **Borders and shading were 100% greenfield.** python-docx has no
  `CT_Border`, `CT_TblBorders`, `CT_TcBorders`, or `CT_Shd` class and
  registers none of those tags. Structurally they are `set_page_borders`
  again, over the `Border` dataclass promoted into `core/` in the
  groundwork PR.
- **Merging already worked.** `_Cell.merge` is fully implemented,
  including the non-rectangular check. `merge_cells` only translates its
  `InvalidSpanError` into a `DocxPlusError` subclass. What was genuinely
  missing is the *inverse*: nothing in python-docx removes a
  `w:gridSpan` or a `w:vMerge`, so a merge was one-way. That is
  `unmerge_cell`.
- **`w:hMerge` was a third thing entirely.** OOXML has two horizontal
  merge encodings and python-docx models only `w:gridSpan`, so a table
  written with `w:hMerge` reads back as separate cells Word draws as
  one. `normalize_horizontal_merges` converts between them.

Verified against Word 2016: the example opens with no repair prompt and
renders the banner span, header shading, inside rules, and the
double-rule callout as intended; an `hMerge` fixture and its normalized
form rasterise byte-for-byte identically. Word's own COM object model
turned out to share python-docx's blind spot, reporting six cells for
the `hMerge` fixture and five after conversion while laying both out the
same way.

Deferred: the **cell-formatting cascade** (table style →
`w:tblStylePr` conditional branch → direct `w:tcPr`).
`read_table_formatting` reports direct formatting only, so a
`Table Grid` table reads back with no borders. That resolver is larger
than every writer in this cycle put together — see the backlog.

### Cross-references to non-bookmark targets — shipped

The backlog called this "mostly instruction grammar over existing
plumbing". That was true for one half and wrong for the other, which is
worth recording:

- **`STYLEREF` was instruction grammar** — it already worked through
  `add_field`. What landed is the typed wrapper
  `fields.add_style_reference`, style-name validation, the switch
  surface, and an outline-level shorthand.
- **Caption references were not.** A `REF` field cannot point at a `SEQ`
  field, only at a bookmark, and `add_caption` created none — so "see
  Figure 3" was not expressible at all, with or without new grammar. The
  fix is `bookmark_name` on `add_caption`, which required promoting
  bookmark emission into `core` (SPEC §9.1 blocks `publishing` from
  importing `bookmarks`) and a bookmark *name* registry, which did not
  exist.

Also filled in the rest of `add_cross_reference`'s switches (`\n` `\r`
`\w` `\p` `\t` `\#` `MERGEFORMAT`), which had only ever supported `\h`,
and added the bookmark-name validation it was missing.

Verified against Word 2016: `REF` to a caption bookmark resolves to
`Figure 1`, `PAGEREF` to `1`, `\p` to `above`, and the header `STYLEREF`
renders `Architecture` on page 1 and `Operations` on page 2.
### Custom numbering — shipped

The largest remaining `python-docx` gap. It exposes a `NumberingPart`
and `len()` of its definitions; there is no `CT_AbstractNum` and no
`CT_Lvl`, so nothing in it can say what a list *looks like*. Landed in
`numbering/`.

Shipped:

- **Define** — `define_list_definition` over `LevelDefinition`, plus
  `define_bullet_list` / `define_numbered_list` presets using Word's own
  glyph and format cycles.
- **Apply** — `apply_list` / `remove_list`, the latter able to write the
  `numId="0"` sentinel that suppresses numbering a *style* applies.
- **Restart** — `restart_list`. OOXML has no paragraph-level "count from
  1 again"; this adds a second `<w:num>` over the same
  `<w:abstractNum>` with a `<w:startOverride>`, as Word does.
- **Read** — `read_list_definitions` returning `ListDefinition` /
  `ListLevel`, tolerant of dangling references and malformed ids.

New plumbing: `NUMBERING_SPEC` (needed because
`DocumentPart.numbering_part` fabricates through an unimplemented stub
that raises), two sequential-allocation registries, and
`assert_numbering_well_formed`.

Verified against Word 2016 via `wordlive` — which caught a cramped
outline where the hanging indent was narrower than the rendered
`1.1.1.`, collapsing the separator tab. Now documented on
`LevelDefinition.hanging`.

Deferred to the backlog: linking a definition into a style definition.

## v0.4 — shipped

### Threaded comments — shipped

The reply / resolve model Word has used since 2013, and the second half
of the collaboration story `revisions/` opened in v0.3. Landed in
`comments/threads.py` over a new `comments/_extended.py`.

Shipped:

- **Replies** — `reply_to_comment` parents a new comment to an existing
  one via `w15:paraIdParent` and mirrors the parent's body-side anchor
  range, which is what makes Word render a thread as a single balloon.
- **Resolve / reopen** — `resolve_comment` / `reopen_comment` toggle
  `w15:done` across the whole thread, matching Word's thread-wide
  Resolve button.
- **Nested read** — `read_threads` returns `CommentThread` (root,
  replies, resolved); `read_comments` results gained `parent_id` and
  `resolved`.
- **Eager metadata** — `add_comment` now stamps `w14:paraId` and writes
  an unresolved `<w15:commentEx>` entry, so every comment is
  thread-ready. `edit_comment` preserves that stamp; `delete_comment`
  cascades to replies by default.
- **CLI** — `docx-plus comments list / resolve / reopen`.

New plumbing: `COMMENTS_EXTENDED_SPEC` (fourth separate part),
`ParaIdRegistry` (the first *package*-wide id namespace), the `w15`
namespace, and a `BUILD_NSMAP` / `NSMAP` split so the extension prefix
stays out of `document.xml`.

Deferred to the backlog: `commentsIds.xml` (`w16cid` durable ids, which
Word regenerates) and `people.xml` (`w15:people` author presence,
cosmetic only).

## v0.3 — shipped

Delivered in the v0.3 cycle.

### 1. Tracked changes (read/write) — shipped (v0.3)

Read/write API for OOXML revision marks, landed in the `revisions/`
module. The canonical "`python-docx` can't do this" gap.

Shipped:

- **Read** — `read_revisions` enumerates every revision type (`w:ins`,
  `w:del`, move wrappers, `w:rPrChange` / `w:pPrChange`, paragraph-mark
  insertions / deletions) with id, author, timestamp, type, and text.
- **Write** — `mark_insertion` / `mark_deletion` wrap existing runs;
  `enable_track_changes` / `disable_track_changes` toggle the
  `settings.xml` flag.
- **Accept / reject** — `accept_revision` / `reject_revision` and the
  `accept_all` / `reject_all` bulk forms resolve insertions and deletions
  fully, with safe non-structural transforms for move and property-change
  marks.

Reused the range/target-normalization pattern from `comments/`, the
`_IdRegistryBase` subclass pattern, and the `settings.xml`-touch pattern
from `fields/update.py`. (Revision marks live inline in `document.xml`, so
no separate part was needed.)

Deferred to the backlog: authoring move pairs and property-change markers
(both need a diff engine), and true paragraph merge/split on accept/reject
of paragraph-mark revisions (currently a non-corrupting fallback).

### 2. CLI — `docx-plus` — shipped (v0.3)

A command-line surface over the existing library, landed in the `cli/`
module and exposed via the `docx-plus` console entry point (also
runnable as `python -m docx_plus.cli`). Built on stdlib `argparse`, so
no new runtime dependency.

Shipped:

- `inspect` — dump effective formatting per paragraph (wraps the cascade
  resolver); `--provenance` and `--json`.
- `restyle` — style remapping (wraps `styles.remap_styles`);
  `--target` / `--map` / `--create-missing`.
- `controls` — `list` / `set` / `clear` content-control values, coercing
  the command-line string to the control's type.

Read commands take `--json`; mutating commands require `-o/--output`
(or an explicit `--in-place`) so the input is never overwritten by
accident.

Still open: a console entry point now exists, so the deferred
**packaging decision for the agent `SKILL.md`** (currently repo-level
only, kept out of the wheel) can be revisited — left for a later cycle.

## Backlog — bounded, unscheduled

Each reuses existing plumbing; pull into a cycle as priority dictates.

- **`commentsExtensible.xml`** — a *fifth* comment side-part
  (`w16cex`, 2018), discovered while verifying the v0.5 work against a
  Word-authored file. Keys off `w16cid:durableId` and carries
  `w16cex:dateUtc`, which exists because `w:comment/@w:date` is local
  time. Low priority: `commentsIds.xml` predates it by two years, so
  writing one without the other is a state Word itself produced for
  years — and Word added no such part when resaving a file of ours that
  lacked it. (`comments/`.)
- **Glossary placeholder text** — the "formal" placeholder mechanism for
  SDTs, vs. the inline `w:placeholder` text `controls/` already supports.
- **Password-protected forms** — legacy hash algorithm, paired with
  `protect_document`. (`protection/`.)
- **linter** -- can use some of the ideas from `wordlive` to build a linter/regularizer for docx files.

## Backlog — larger or dependency-gated

- **Custom XML Parts data binding** — wires repeating-section content
  controls to a custom XML data source: new relationship types and
  `<w:dataBinding>` children on SDTs. `core/parts.py` already supports
  separate parts. **Gates the next item.**
- **Bibliography** — sources in a Custom XML Part, `<w:sdt>` citations
  referencing them, a `BIBLIOGRAPHY` field rendering the list. Rides on
  the data-binding subsystem above.
- **Theme writing** — `styles/theme.py` reads themes today; writing rounds
  out the surface.
- **High-level "restyle" planner** — inverse of the inspector: take a
  target `ResolvedFormatting` and compute the minimal cascade modification
  to reach it. Large design space.
- **Sections / headers / footers first-class API** — wraps the
  `python-docx` primitives behind a `docx_plus`-native surface
  (`sections/`).
- **Cell-formatting cascade resolver** — resolve a cell's effective
  borders, shading, and margins through table style → `w:tblStylePr`
  conditional branch (first row / last column / banding) → direct
  `w:tcPr`. `tables/read.py` reports direct formatting only, and
  `styles/inspect.py` scopes this out in the same terms while resolving
  the paragraph and run cascade. The largest single item on this list.
- **Link a numbering definition into a style** —
  `w:style/w:pPr/w:numPr`, so `ensure_style("ListBullet")` produces a
  style that actually bullets. Split out of the v0.5 numbering cycle
  because `styles/modify.py` already owns writing into `w:style` and
  carries `_STYLE_CHILD_ORDER` / `_PPR_CHILD_ORDER`; duplicating that in
  `numbering/` would put the same schema knowledge in two places.
  (`styles/`.)

  Concretely, so this doesn't need re-deriving:

  - The seam is documented in the code already —
    `styles/modify.py` `_BUILTIN_STYLES` Tier D carries a comment saying
    the `List*` styles omit `numPr` and that "callers wanting actual
    auto-numbering should attach a numbering definition separately".
    Asymmetry to fix: the *bundled template's* `ListBullet` does link
    `numId` 1, so `ensure_style` on a document lacking the style
    produces something weaker than a stock `Document()` already has.
  - The work is a new key in `_write_paragraph_property`
    (`modify.py:835`) plus an entry in `_validate_property_keys`
    (`:818`), which currently rejects anything outside the
    paragraph/run property sets. `_PPR_CHILD_ORDER` already reserves the
    `numPr` slot; `_ordered_insert` places it.
  - No cross-capability import is needed: linking is just writing
    `w:numPr` with an int, so `styles/` never has to reach into
    `numbering/`. That is why this lands in `styles/`, not here.
- **Resolve style-supplied numbering in the cascade** — related, and
  arguably the more surprising half. Layer 4 reads only the paragraph's
  *direct* `w:numPr` (`styles/inspect.py:428`), never the one its style
  supplies, so on a stock `Document()`:

  ```python
  p = doc.add_paragraph("bulleted", style="List Bullet")
  resolve_effective_formatting(p).num_id   # -> None
  ```

  even though that style links `numId` 1 in the bundled template
  (verified 2026-07-26). Every other property on `ResolvedFormatting`
  walks the style chain, so `num_id` silently breaks the contract the
  rest of the dataclass sets. Fixing it means resolving `numPr` through
  the `basedOn` chain like the other pPr properties, and deciding what
  `numId` `0` from a style should mean (the "no numbering" sentinel).
  Wants its own tests; do it alongside or before the style-linking item
  above, since that item makes the gap much easier to hit. (`styles/`.)

## Considered, not on the roadmap

- **MCP server** — surfaced in the old `notes.md` scratchpad alongside the
  CLI. An MCP wrapper is an adjacent product, not a `python-docx`
  extension; route to a sibling project rather than absorbing it here.
  Revisit only if the CLI lands and a thin MCP front end over it is
  cheap.
