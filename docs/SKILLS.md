# Agent skill (LLM usage)

`docx_plus` ships an **Agent Skill** — an LLM-facing guide that teaches a
coding agent how to drive the library correctly: the units, the
"`mark_fields_dirty` before save" rule for fields, style IDs vs. names, toggle
semantics, and the per-module APIs. If you use [Claude
Code](https://claude.com/claude-code) (or any agent that reads skill files) to
generate `.docx` automation, point it at this skill instead of hand-feeding API
snippets.

"Ships" is literal: since v0.5 the skill lives at `docx_plus/skill/` **inside
the package**, so `pip install docx-plus` puts it on disk. You do not need to
clone the repository to get it.

## Installing it

```bash
docx-plus skill install            # -> ./.claude/skills/docx-plus
docx-plus skill install --user     # -> ~/.claude/skills/docx-plus
docx-plus skill install --dest DIR # -> DIR/docx-plus
```

An existing installation is left alone unless you pass `--force`, so upgrading
the library never silently clobbers local edits — re-run with `--force` to pick
up a new release's version.

Claude Code discovers the skill through the `name:` / `description:`
frontmatter in `SKILL.md` and loads it when a task matches.

## Reading it without installing

```bash
docx-plus skill path               # where the packaged copy lives
docx-plus skill list               # the reference topics
docx-plus skill show               # print SKILL.md
docx-plus skill show tables        # print one topic
```

For any other LLM / RAG pipeline, feed `SKILL.md` as context; it points at the
reference files, which the agent reads on demand. Everything is plain Markdown
with runnable Python.

## What's in it

The standard umbrella + progressive-disclosure layout: one `SKILL.md` entry
point an agent loads first, plus topic reference files it pulls in on demand.

| File | Covers |
|---|---|
| `SKILL.md` | Entry point — mental model, cross-cutting conventions, capability map |
| `reference/forms.md` | `FormBuilder`; read / set / clear control values; document protection |
| `reference/styles.md` | Cascade inspection + provenance; create / modify / apply / ensure / remap; theme |
| `reference/publishing.md` | TOC, captions, table of figures, footnotes, endnotes, bookmarks, cross-references, fields |
| `reference/layout.md` | Columns, section breaks, even/odd headers, line numbering, page borders |
| `reference/numbering.md` | Custom bullet / numbered / multilevel list definitions, applying and restarting |
| `reference/tables.md` | Table and cell borders, shading, merging / unmerging, `w:hMerge` normalization |
| `reference/comments.md` | Anchored comments, threads, durable ids, author presence |
| `reference/revisions.md` | Tracked changes (mark / read / accept / reject); track-changes toggle |
| `reference/lint.md` | Audit a document for formatting defects; describe the repair as a plan. Read-only |
| `reference/cli.md` | The `docx-plus` command line |

## Accuracy

Every code snippet in the skill imports only public symbols and is verified to
run end to end against the current release (v0.6.0). The skill mirrors the same
public surface documented in the [API Index](API.md) and the per-module
[Reference](reference/core-ns.md) pages — it's the agent-facing complement to
those human-facing docs, not a separate source of truth.

The test suite asserts that every topic file is reachable from `SKILL.md` and
that the frontmatter carries the `name:` / `description:` an agent needs to
discover it, so a new reference page cannot land orphaned.
