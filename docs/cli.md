# Command line

Installing `docx-plus` puts a `docx-plus` console command on your `PATH` (it is
also runnable as `python -m docx_plus.cli`). The CLI is a thin shell over the
library — each subcommand wraps one tested function — for the cases where you
want to inspect or edit a `.docx` from a shell or a script without writing
Python.

```console
$ docx-plus --help
usage: docx-plus [-h] [--version] <command> ...

  inspect   dump the effective formatting of each paragraph
  restyle   remap a document's styles onto canonical style ids
  controls  list, set, or clear content-control values
```

Two conventions hold across every command:

- **Read commands take `--json`.** `inspect` and `controls list` default to
  human-readable text; `--json` emits structured output for piping into `jq` or
  another tool.
- **Mutating commands never overwrite the input by accident.** `restyle` and
  `controls set` / `clear` require an explicit `-o/--output` path, or the
  `--in-place` flag to opt into overwriting the source file.

Exit codes: `0` on success, `1` for a handled error (bad path, missing output,
un-coercible value, unknown control tag), `2` for a usage error or when no
command is given.

## `inspect`

Resolve the effective formatting for every paragraph, wrapping
[`resolve_effective_formatting`](reference/styles-inspect.md).

```console
$ docx-plus inspect report.docx
[1] "Quarterly Review"
    style: Title
    font_name : 'Calibri Light'
    font_size : 28.0
    bold      : True

$ docx-plus inspect report.docx --provenance
[1] "Quarterly Review"
    style: Title
    font_size : 28.0   <- paragraphStyle: Title
    bold      : True   <- paragraphStyle: Title (toggle XOR)
```

`--provenance` annotates each field with the cascade layer (and style id) that
set it. `--json` emits one record per paragraph with `index`, `text`,
`style_id`, `style_name`, `partial`, a `fields` object, and — when
`--provenance` is set — a `provenance` object.

## `restyle`

Reconcile a document's styles against a set of canonical ids, wrapping
[`remap_styles`](reference/styles-modify.md). Paragraphs and runs are remapped
onto the resolved styles and the `target -> resolved-id` mapping is reported.

```console
$ docx-plus restyle draft.docx --target Heading1 --target Title -o clean.docx
wrote clean.docx
  Heading1 -> Heading1
  Title    -> Title
```

- `--target STYLE_ID` (repeatable, required) — the canonical ids to reconcile.
- `--map TARGET=EXISTING` (repeatable) — hint resolving a target to a specific
  existing style id.
- `--create-missing` — materialize known built-in targets that aren't defined
  in the document yet.
- `-o/--output` (or `--in-place`) — where to write the result.
- `--json` — emit the resolved mapping as JSON.

## `controls`

List, set, or clear content controls (fillable form fields), wrapping
[`read_controls` / `set_control_value` / `clear_control`](reference/controls-read.md).

```console
$ docx-plus controls list form.docx
name: text alias='Full name' = (placeholder)
dept: dropdown = 'Engineering'
subscribed: checkbox = False

$ docx-plus controls set form.docx --tag name --value "Ada Lovelace" -o filled.docx
set 'name' = 'Ada Lovelace'; wrote filled.docx

$ docx-plus controls clear filled.docx --tag name --in-place
cleared 'name'; wrote filled.docx
```

- `controls list FILE [--by tag|alias] [--json]` — every control with its tag,
  alias, type, value, and placeholder state. `--by alias` keys on the alias and
  skips controls without one.
- `controls set FILE --tag T --value V -o OUT` — the command reads the control's
  type and coerces the string `V`: `true/false/1/0/yes/no` for checkboxes, an
  ISO 8601 string (`2026-06-15`) for dates, plain text otherwise. An un-coercible
  value or unknown tag is a clean error.
- `controls clear FILE --tag T -o OUT` — reset the control to its placeholder
  state.
