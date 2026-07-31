# CLI — the `docx-plus` console command

Module: `docx_plus.cli`. A command-line surface over the library, installed as
the `docx-plus` console command (also runnable as `python -m docx_plus.cli`).

**Why this exists:** the library is Python-first, but inspecting or batch-editing
a `.docx` from a shell or a Makefile shouldn't require writing a script. Each
subcommand is a thin wrapper over one tested function. Built on stdlib
`argparse` — no new runtime dependency.

**You usually don't need this from Python.** If you're already writing Python,
call the underlying functions directly (`resolve_effective_formatting`,
`remap_styles`, `read_controls` / `set_control_value` / `clear_control`,
`read_threads` / `resolve_comment`, `lint` / `plan_fixes`). Reach for the CLI
when the caller is a shell, CI step, or another non-Python tool.

## Conventions

- **Read commands take `--json`** (`inspect`, `controls list`,
  `comments list`, `lint`, `plan`): default is human-readable text; `--json`
  emits structured output.
- **Mutating commands never overwrite the input by accident** (`restyle`,
  `controls set` / `clear`, `comments resolve` / `reopen`): they require
  `-o/--output`, or `--in-place` to opt into overwriting the source.
- **Exit codes:** `0` success; `1` for a handled error (bad path, missing
  output, un-coercible value, unknown tag or comment id), printed as a
  single line on stderr; `2` for a usage error or no command. A mistake the
  CLI catches is prefixed `error:`; a typed library error names its class
  instead (`UnknownRuleError: ...`, `InvalidProfileError: ...`).
- **`lint` and `plan` overload exit `1`** as "I found something", so they
  gate a CI step directly. Both are read-only, so there is nothing to
  overwrite.
- **`skill` is the exception to all of the above** — it touches no `.docx`,
  so it takes neither `--json` nor `-o/--output`.

## `inspect` — effective formatting

```bash
docx-plus inspect report.docx                 # text, per paragraph
docx-plus inspect report.docx --provenance    # annotate each field with its cascade layer
docx-plus inspect report.docx --json          # structured records
```

Wraps `resolve_effective_formatting`. JSON record per paragraph: `index`,
`text`, `style_id`, `style_name`, `partial`, `fields` (only the cascade-set
fields), and — with `--provenance` — a `provenance` map of `field -> layer`.
The `index` is **1-based** (paragraphs are numbered from 1), which differs from
the 0-based `paragraph_index` used by the library's `read_*` functions
(`read_revisions`, `read_comments`, etc.).

## `restyle` — style remapping

```bash
docx-plus restyle draft.docx --target Heading1 --target Title -o clean.docx
docx-plus restyle draft.docx --target Heading1 --map Heading1=Heading2 -o clean.docx
docx-plus restyle draft.docx --target Quote --create-missing -o clean.docx
```

Wraps `remap_styles`. `--target STYLE_ID` is repeatable and required; `--map
TARGET=EXISTING` is a repeatable resolution hint; `--create-missing`
materializes known built-in targets. Reports the `target -> resolved-id`
mapping (text, or `--json`). Output goes to `-o/--output` (or `--in-place`).

## `controls` — list / set / clear

```bash
docx-plus controls list form.docx                       # tag: type [alias=...] = value
docx-plus controls list form.docx --by alias --json
docx-plus controls set form.docx --tag name --value "Ada Lovelace" -o filled.docx
docx-plus controls set form.docx --tag subscribed --value true -o filled.docx
docx-plus controls set form.docx --tag start --value 2026-06-15 -o filled.docx
docx-plus controls clear filled.docx --tag name --in-place
```

Wraps `read_controls` / `set_control_value` / `clear_control`. `set` reads the
control's type and **coerces the command-line string**: `true/false/1/0/yes/no/on/off`
for checkboxes, an ISO 8601 string for dates, plain text otherwise. An
un-coercible value or unknown tag is a clean `error: ...` (exit 1).

## `comments` — list / resolve / reopen

```bash
docx-plus comments list draft.docx                      # threads, replies indented
docx-plus comments list draft.docx --unresolved --json  # only open threads
docx-plus comments resolve draft.docx 3 -o triaged.docx
docx-plus comments reopen triaged.docx 3 --in-place
```

Wraps `read_threads` / `resolve_comment` / `reopen_comment`. `list` prints each
thread root with its replies indented beneath, the anchored text each is
attached to, and a `[resolved]` marker; a comment with no anchor in the document
body is flagged as orphaned. Resolution is **thread-wide**, so `resolve` /
`reopen` accept any comment id in the thread. An unknown id is a clean
`error: ...` (exit 1).

## `lint` — audit formatting

```bash
docx-plus lint report.docx                      # findings, human-readable
docx-plus lint report.docx --json               # the full finding shape
docx-plus lint report.docx --rule typography    # one cluster of rules
docx-plus lint report.docx --exclude double-space
docx-plus lint --list-rules                     # the catalogue; no document needed
docx-plus lint report.docx --profile house.json
```

Wraps `docx_plus.lint.lint`. Read-only. Output is one finding per block —
severity mark, location, rule id, message, and the paragraph excerpt quoted
beneath:

```
W paragraph 1         heading-level-skip           Outline jumps from level 1 to level 3, skipping level 2.
                                                   > "Deep Dive"
W paragraph 3         manual-list                  Paragraph begins with a typed list marker but carries no numbering.
                                                   > "1. First typed item"
W paragraph 5         style-drift                  Paragraph overrides space after directly, deviating from Normal.
                                                   > "Drifting paragraph"
i paragraph 2         double-space                 Two or more consecutive spaces between words.
                                                   > "Body with  two spaces and a space ."
i paragraph 2         space-before-punctuation     Whitespace before '.'.
                                                   > "Body with  two spaces and a space ."
i paragraph 4, run 0  redundant-direct-formatting  Run sets size directly to the value it already inherits; the direct formatting has no effect but overrides the style.
                                                   > "redundant"

6 findings (3 warning, 3 info).
```

Excerpts are quoted because several rules are about whitespace nobody can see;
a tab prints as `\t`.

`--rule` and `--exclude` both take a rule **id** or a **tag**, and both repeat.
Naming a tag also *enables* that cluster's off-by-default rules. A selector
matching nothing is an error, not a silent empty result. `--no-tables` skips
paragraphs inside table cells; headers, footers, notes, and comments are never
audited.

**Exit code `1` when it found anything**, so `docx-plus lint doc.docx` works as
a CI gate on its own. See `reference/lint.md` for the rule catalogue, the
`consistency` / `structural` / `policy` kinds, and profiles.

## `plan` — describe the repair

```bash
docx-plus plan report.docx                    # the edits, in applying order
docx-plus plan report.docx --allow-content    # include deletions
docx-plus plan report.docx --json             # the whole plan, serialized
```

Wraps `plan_fixes`. Also read-only — **this release applies nothing.** The
command exists so a repair is inspectable before anything can perform it.

```
4 edit(s), in the order they would be applied:

1. [check] paragraph 2  double-space
     Collapse 1 run of spaces to a single space.
     - replace-paragraph-text paragraph_index=2 spans=[9-11->' ']
2. [check] paragraph 2  space-before-punctuation
     Remove the whitespace before 1 punctuation mark.
     - replace-paragraph-text paragraph_index=2 spans=[33-34->'']
3. [safe ] paragraph 4, run 0  redundant-direct-formatting
     Delete the run's direct size; the style supplies the same value.
     - clear-run-properties paragraph_index=4 run_index=0 properties=font_size
4. [check] paragraph 5  style-drift
     Clear the paragraph's direct space after so Normal applies.
     - clear-paragraph-properties paragraph_index=5 properties=spacing_after

2 finding(s) with no known repair:
     1  heading-level-skip
     1  manual-list

4 to apply, 0 withheld, 0 dropped, 2 unfixable.
```

Four possible sections, for the four things that can happen to a finding: it
becomes an edit, it is **withheld** for changing what the document contains
(deleting a paragraph or a style — `--allow-content` includes those), it is
**dropped** for colliding with an earlier edit, or nobody knows how to fix it.
Every finding lands in exactly one, so the report accounts for the whole audit.

The bracketed mark is the safety class: `safe` renders identically afterwards,
`check` changes rendering or text deliberately, `DROP` removes something that
cannot be reconstructed.

Exit code `1` when the plan holds any edit, applied or withheld. `--rule`,
`--exclude`, `--no-tables`, `--profile`, and `--no-profile` work exactly as
they do for `lint`.

## `skill` — the packaged agent skill

```bash
docx-plus skill install                  # -> ./.claude/skills/docx-plus
docx-plus skill install --user           # -> ~/.claude/skills/docx-plus
docx-plus skill install --dest DIR       # -> DIR/docx-plus
docx-plus skill path                     # where the packaged copy lives
docx-plus skill list                     # the reference topics
docx-plus skill show tables              # print one topic to stdout
```

This is the file you are reading. It ships **inside the wheel** at
`docx_plus/skill/`, so `pip install docx-plus` is enough — no repository
clone. `install` refuses to overwrite an existing installation without
`--force`, so upgrading the library never silently clobbers local edits.
`--user` and `--dest` are mutually exclusive.
