# docx_plus — Roadmap

The single authoritative roadmap for `docx_plus`. `SPEC.md §15` (a v0.1-era
historical list) and `docs/ARCHITECTURE.md §11` both defer to this file.

`docx_plus` is, and stays, a lean extension to `python-docx` that does the
things `python-docx` can't. Every item below either fills a documented
`python-docx` gap or rounds out a surface already started here. Ideas that
don't fit that charter are routed to sibling projects, not absorbed.

## Current state — v0.2.1 (released 2026-05-21)

Tagged: `v0.1.0`, `v0.2.0`, `v0.2.1`. Shipped capability modules:

| Module | Surface |
|---|---|
| `styles/` | Cascade inspection (`ResolvedFormatting`, all 12 toggles, theme fonts + colors, conditional table styles), modification, remapping |
| `controls/` | Content controls — `FormBuilder`, read / set / clear values |
| `fields/` | Simple + complex fields, `mark_fields_dirty` |
| `protection/` | `protect_document` |
| `comments/` | Anchored comments — add / edit / delete / clear, over runs, paragraphs, run ranges |
| `layout/` | Columns, mid-document section breaks, even/odd headers, line numbering, page borders |
| `bookmarks/` | Bookmarks + `REF` / `PAGEREF` cross-references |
| `notes/` | Footnotes + endnotes — add / edit / read |
| `publishing/` | TOC, captions, table of figures |

Suite at last gate: 717 tests (709 pass, 8 LibreOffice-skipped); `mypy
--strict`, `ruff`, and `mkdocs build --strict` all clean.

## v0.3 — targeted next

Prioritized for the next cycle.

### 1. Tracked changes (read/write)

Read/write API for OOXML revision marks — `w:ins`, `w:del`,
`w:moveFromRangeStart` / `w:moveToRangeStart`, and friends. The canonical
"`python-docx` can't do this" gap and the highest-value remaining surface.
Largest single scope on the roadmap; expect it to anchor the cycle.

- New `revisions/` (or `changes/`) module.
- Read first (enumerate revisions, author, timestamp, anchored text),
  then write (wrap insertions/deletions, accept/reject).
- Reuses the separate-parts and range-anchoring patterns proven in
  `comments/`.

### 2. CLI — `docx-plus`

A command-line surface over the existing library:

- `restyle` — style remapping (wraps `styles.remap_styles`).
- `inspect` — dump effective formatting (wraps the cascade resolver).
- `controls` — list / set content-control values.

Landing a core CLI also reopens the **packaging decision for the agent
`SKILL.md`** (currently repo-level only, deliberately kept out of the
wheel) — revisit bundling it once a console entry point exists.

## v0.3+ backlog — bounded, unscheduled

Each reuses existing plumbing; pull into a cycle as priority dictates.

- **Cross-references to non-bookmark targets** — `STYLEREF` for
  heading-text references and sequence (`SEQ`) fields for caption / figure
  numbering. Reuses the complex-field plumbing; the work is the
  instruction grammar. (`bookmarks/` or a new `crossref/`.)
- **Threaded comments + resolve / reopen** — `w15 parentCommentEx` for
  parent/child replies plus the respond / resolve / reopen ops. Adds a
  `w15` namespace dependency and a separate `commentsExtended.xml` part.
  Completes the comments story. (`comments/`.)
- **Glossary placeholder text** — the "formal" placeholder mechanism for
  SDTs, vs. the inline `w:placeholder` text `controls/` already supports.
- **Password-protected forms** — legacy hash algorithm, paired with
  `protect_document`. (`protection/`.)

## v0.3+ backlog — larger or dependency-gated

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
- **Custom numbering definitions** — a `numbering/` module.

## Considered, not on the roadmap

- **MCP server** — surfaced in the old `notes.md` scratchpad alongside the
  CLI. An MCP wrapper is an adjacent product, not a `python-docx`
  extension; route to a sibling project rather than absorbing it here.
  Revisit only if the CLI lands and a thin MCP front end over it is
  cheap.
