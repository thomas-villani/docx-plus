# Agent skill (LLM usage)

`docx_plus` ships an **Agent Skill** — an LLM-facing guide that teaches a
coding agent how to drive the library correctly: the units, the
"`mark_fields_dirty` before save" rule for fields, style IDs vs. names, toggle
semantics, and the per-module APIs. If you use [Claude
Code](https://claude.com/claude-code) (or any agent that reads skill files) to
generate `.docx` automation, point it at this skill instead of hand-feeding API
snippets.

## What's in it

The skill lives at
[`skills/docx-plus/`](https://github.com/thomas-villani/docx-plus/tree/main/skills/docx-plus)
in the repository, using the standard umbrella + progressive-disclosure layout:
one `SKILL.md` entry point that an agent loads first, plus topic reference files
it pulls in on demand.

| File | Covers |
|---|---|
| [`SKILL.md`](https://github.com/thomas-villani/docx-plus/blob/main/skills/docx-plus/SKILL.md) | Entry point — mental model, cross-cutting conventions, capability map |
| [`reference/forms.md`](https://github.com/thomas-villani/docx-plus/blob/main/skills/docx-plus/reference/forms.md) | `FormBuilder`; read / set / clear control values; document protection |
| [`reference/styles.md`](https://github.com/thomas-villani/docx-plus/blob/main/skills/docx-plus/reference/styles.md) | Cascade inspection + provenance; create / modify / apply / ensure / remap; theme |
| [`reference/publishing.md`](https://github.com/thomas-villani/docx-plus/blob/main/skills/docx-plus/reference/publishing.md) | TOC, captions, table of figures, footnotes, endnotes, bookmarks, cross-references, fields |
| [`reference/layout.md`](https://github.com/thomas-villani/docx-plus/blob/main/skills/docx-plus/reference/layout.md) | Columns, section breaks, even/odd headers, line numbering, page borders |
| [`reference/comments.md`](https://github.com/thomas-villani/docx-plus/blob/main/skills/docx-plus/reference/comments.md) | Anchored comments (add / edit / delete / read) |

## Using it

- **Claude Code** — copy (or symlink) the `skills/docx-plus` directory into your
  project's `.claude/skills/` folder, or into your personal `~/.claude/skills/`.
  The agent discovers it through the `name:` / `description:` frontmatter and
  loads it when a task matches.
- **Any other LLM / RAG pipeline** — feed `SKILL.md` as context; it points to
  the reference files, which the agent reads on demand. Everything is plain
  Markdown with runnable Python.

## Accuracy

Every code snippet in the skill imports only public symbols and is verified to
run end to end against the current release (v0.2.0). The skill mirrors the same
public surface documented in the [API Index](API.md) and the per-module
[Reference](reference/core-ns.md) pages — it's the agent-facing complement to
those human-facing docs, not a separate source of truth.
