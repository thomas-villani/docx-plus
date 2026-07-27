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
| `comments/` | Anchored comments — add / edit / delete / clear, over runs, paragraphs, run ranges; threads — reply, resolve / reopen, nested read (v0.4) |
| `layout/` | Columns, mid-document section breaks, even/odd headers, line numbering, page borders |
| `bookmarks/` | Bookmarks + `REF` / `PAGEREF` cross-references |
| `notes/` | Footnotes + endnotes — add / edit / read |
| `publishing/` | TOC, captions, table of figures |
| `revisions/` | Tracked changes — mark insertions / deletions, read revisions, accept / reject, track-changes toggle (v0.3) |
| `cli/` | `docx-plus` console command — `inspect` (effective formatting), `restyle` (style remapping), `controls` (list / set / clear values) (v0.3), `comments` (list / resolve / reopen threads) (v0.4) |

Suite at the v0.4.0 release: 905 tests (895 pass, 10 LibreOffice-skipped),
94% coverage; `mypy --strict`, `ruff`, and `mkdocs build --strict` all
clean.

## v0.5 — in progress

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

- **Cross-references to non-bookmark targets** — `STYLEREF` for
  heading-text references, plus references *to* an existing caption
  ("see Figure 3"). Smaller than it first looks: `SEQ` authoring already
  shipped in `publishing/captions.py`, so what is left is the instruction
  grammar over the existing complex-field plumbing. (`bookmarks/` or a
  new `crossref/`.)
- **Comment durable ids + author presence** — `commentsIds.xml`
  (`w16cid` durable ids for comment permalinks) and `people.xml`
  (`w15:people` presence info). Neither is needed for threading or
  resolution — Word regenerates the first and the second is cosmetic —
  so both were split out of the v0.4 cycle. (`comments/`.)
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
- **Table cell merging / borders / shading** beyond `python-docx`
  defaults. (Distinct from *page* borders, already shipped in
  `layout/borders.py`.)
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
