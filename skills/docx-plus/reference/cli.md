# CLI — the `docx-plus` console command

Module: `docx_plus.cli`. A command-line surface over the library, installed as
the `docx-plus` console command (also runnable as `python -m docx_plus.cli`).

**Why this exists:** the library is Python-first, but inspecting or batch-editing
a `.docx` from a shell or a Makefile shouldn't require writing a script. Each
subcommand is a thin wrapper over one tested function. Built on stdlib
`argparse` — no new runtime dependency.

**You usually don't need this from Python.** If you're already writing Python,
call the underlying functions directly (`resolve_effective_formatting`,
`remap_styles`, `read_controls` / `set_control_value` / `clear_control`). Reach
for the CLI when the caller is a shell, CI step, or another non-Python tool.

## Conventions

- **Read commands take `--json`** (`inspect`, `controls list`): default is
  human-readable text; `--json` emits structured output.
- **Mutating commands never overwrite the input by accident** (`restyle`,
  `controls set` / `clear`): they require `-o/--output`, or `--in-place` to
  opt into overwriting the source.
- **Exit codes:** `0` success; `1` for a handled error (bad path, missing
  output, un-coercible value, unknown tag — printed as `error: ...` on stderr);
  `2` for a usage error or no command.

## `inspect` — effective formatting

```bash
docx-plus inspect report.docx                 # text, per paragraph
docx-plus inspect report.docx --provenance    # annotate each field with its cascade layer
docx-plus inspect report.docx --json          # structured records
```

Wraps `resolve_effective_formatting`. JSON record per paragraph: `index`,
`text`, `style_id`, `style_name`, `partial`, `fields` (only the cascade-set
fields), and — with `--provenance` — a `provenance` map of `field -> layer`.

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
control's type and **coerces the command-line string**: `true/false/1/0/yes/no`
for checkboxes, an ISO 8601 string for dates, plain text otherwise. An
un-coercible value or unknown tag is a clean `error: ...` (exit 1).
