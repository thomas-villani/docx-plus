# Concepts

Present-tense reference for how `docx_plus` is laid out and **why**. These
pages describe what currently exists at the end of the v0.6 cycle.

They are the *explanation* half of the documentation. If you are trying to
get something done, start with the [guides](../guides/index.md) — each one
links back here for the reasoning. If you want a signature, go to the
[API index](../API.md) or the [module reference](../reference/core-ns.md).

**Audience:** a developer extending or debugging `docx_plus` itself, or a
user who wants more than the guides before reading source.

The contract that constrains the library is `SPEC.md`; the meta-guidance on
how it was built and how to extend it is `IMPLEMENTATION.md`. Read these
pages when you need to understand the library's shape; read those when you
need to decide what to add, or how.

## The pages

### Foundations

| Page | What it covers |
|---|---|
| [Package layout](package-layout.md) | The directory tree, and why it is flat |
| [Schema-strict insertion](schema-order.md) | Why every write goes through an ordered insert |
| [Separate OOXML parts](parts.md) | `comments.xml`, `footnotes.xml`, `numbering.xml`, and the five side-parts |
| [Invariants, errors, and testing](invariants.md) | The eight architectural commitments, the error hierarchy, and the three test layers |

### The cascade

| Page | What it covers |
|---|---|
| [The cascade resolver](cascade.md) | Six layers, toggle properties, paragraph spacing, theme colours, provenance |
| [Styles: remapping and built-ins](styles.md) | Style ids vs. names, `remap_styles`, and the 107-entry built-in table |

### Capabilities

| Page | What it covers |
|---|---|
| [Content controls](controls.md) | The five SDT types, and the read / write split |
| [Fields and protection](fields.md) | The five-run complex field, `mark_fields_dirty`, `documentProtection` |
| [Anchored comments](comments.md) | The five elements per comment, threading, durable ids, author presence |
| [Layout](layout.md) | Columns, mid-document section breaks, even/odd headers, line numbering, page borders |
| [Bookmarks and cross-references](bookmarks.md) | Paired markers, `REF` / `PAGEREF` |
| [Footnotes and endnotes](notes.md) | The two separate parts, and the reserved ids |
| [Publishing](publishing.md) | TOC, `SEQ` captions, table of figures |
| [Tracked changes](revisions.md) | `w:ins` / `w:del`, accept / reject, track-changes mode |
| [Custom numbering](numbering.md) | The abstract/instance model, and the indent trap |
| [Table formatting](tables.md) | Borders, shading, and the two horizontal-merge encodings |
| [Linting and the fix plan](lint.md) | Rule kinds, the fix vocabulary, what only the planner can decide |
| [The CLI](cli.md) | The composition layer, and the packaged agent skill |

## What's next

v0.1 (Phases 1–6), the v0.2 cycle, and the v0.2 in-place expansion are
complete (released through `v0.2.1`). v0.3 then shipped its two headline
targets: **tracked changes (read/write)** in `revisions/` and the
**`docx-plus` CLI** in `cli/`. v0.4 shipped **threaded comments** in
`comments/threads.py` over `comments/_extended.py`, with the
`commentsExtended.xml` part and the `comments` CLI subcommand. v0.5
shipped **custom numbering** in `numbering/`, **table formatting** in
`tables/`, comment durable ids and author presence, `STYLEREF` and
caption cross-references, and moved the agent skill into the wheel
behind a `docx-plus skill` command.

v0.6 is scoped to the **linter** — a new `lint/` composing layer that sits
where `cli/` sits, above the capability modules, reporting findings over
the style cascade and producing an inspectable fix plan. Both halves have
landed: twenty rules behind `docx-plus lint`, and
`plan_fixes(findings) -> FixPlan` behind `docx-plus plan`. The cycle is
non-mutating throughout — a plan is a serializable description of edits
and nothing applies one; that is v0.7.

The authoritative roadmap for that cycle and for everything on the
backlog — bounded items and dependency-gated ones (the cell-formatting
cascade resolver, bibliography / CXML data binding, theme writing, …) —
lives in
[`ROADMAP.md`](https://github.com/thomas-villani/docx-plus/blob/main/ROADMAP.md)
at the repo root.
