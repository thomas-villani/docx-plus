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
  comments  list comment threads, or resolve / reopen them
  lint      audit a document for formatting defects
  skill     locate, read, or install the packaged agent skill
```

Two conventions hold across every command that touches a `.docx`:

- **Read commands take `--json`.** `inspect`, `controls list`, and
  `comments list` default to human-readable text; `--json` emits structured
  output for piping into `jq` or another tool.
- **Mutating commands never overwrite the input by accident.** `restyle`,
  `controls set` / `clear`, and `comments resolve` / `reopen` require an
  explicit `-o/--output` path, or the `--in-place` flag to opt into
  overwriting the source file.

`skill` is the exception to both: it neither reads nor writes a `.docx`, so it
takes neither `--json` nor `-o/--output`.

Exit codes: `0` on success, `1` for a handled error (bad path, missing output,
un-coercible value, unknown control tag, unknown comment id), `2` for a usage
error or when no command is given.

`lint` overloads `1` deliberately: a successful audit that *found something*
also exits `1`, so the command works as a CI gate. A genuine failure is still
distinguishable — it prints an `error:` line to stderr and no findings.

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

## `comments`

List comment threads, or triage them by resolving and reopening, wrapping
[`read_threads` / `resolve_comment` / `reopen_comment`](reference/comments-threads.md).

```console
$ docx-plus comments list draft.docx
[1] Alice [resolved]: Where does the six-week figure come from?
    on paragraph 1: 'a six-week schedule'
  [2] Bob: From the Q2 capacity model.
      on paragraph 1: 'a six-week schedule'
[3] Carol: These need refreshing before we circulate.
    on paragraph 2: 'Budget figures are carried over from last quarter.'

$ docx-plus comments list draft.docx --unresolved --json | jq '.[].author'
"Carol"

$ docx-plus comments resolve draft.docx 3 -o triaged.docx
resolved the thread containing comment 3; wrote triaged.docx
```

- `comments list FILE [--unresolved] [--json]` — every thread, replies indented
  under their root, with the anchored text each is attached to. A comment with
  no anchor in the document body is flagged as orphaned. `--unresolved` hides
  threads that are already closed.
- `comments resolve FILE ID -o OUT` / `comments reopen FILE ID -o OUT` — toggle
  a thread's resolved state. Resolution is thread-wide in Word, so naming any
  comment in a thread moves the whole thread; an unknown id is a clean error.

## `lint`

Audit a document for formatting defects, wrapping
[`docx_plus.lint.lint`](reference/lint.md). Read-only: it reports, and
changes nothing.

```console
$ docx-plus lint report.docx
W paragraph 1         heading-level-skip           Outline jumps from level 1 to level 3, skipping level 2.
                                                   > "Deep Dive"
W paragraph 3         manual-list                  Paragraph begins with a typed list marker but carries no numbering.
                                                   > "1. First typed item"
i paragraph 4         double-space                 Two or more consecutive spaces between words.
                                                   > "Body with  two spaces and a space ."
i paragraph 9, run 0  redundant-direct-formatting  Run sets size directly to the value it already inherits.
                                                   > "redundant"

8 findings (4 warning, 4 info).
```

Excerpts are quoted because several rules are about whitespace nobody can
see; a tab prints as `\t`.

**Selecting rules.** `--rule` and `--exclude` both take a rule **id** or a
**tag**, and both repeat:

```console
$ docx-plus lint report.docx --rule typography      # one cluster
$ docx-plus lint report.docx --exclude double-space # everything but one rule
$ docx-plus lint report.docx --rule whitespace --rule structure
```

Naming a tag also **enables** that cluster's off-by-default rules — that is
how you opt into the heuristic ones without listing them individually. A
selector matching no rule or tag is an error rather than a silent empty
result, since "no findings" and "no rules ran" look identical otherwise.

**The catalogue.** `--list-rules` needs no document:

```console
$ docx-plus lint --list-rules
double-space                 consistency  info     on   [typography,whitespace]
                             Two or more spaces between words in body text.
heading-level-skip           structural   warning  on   [headings,structure]
                             The outline jumps a level (e.g. Heading 1 straight to Heading 3).
mixed-run-formatting         consistency  info     off  [formatting]
                             Runs within one paragraph disagree on font or size.
...
10 rules, 8 on by default.
```

The `consistency` / `structural` / `policy` column is the rule's **kind**,
and it is worth understanding before you act on output:

- **`consistency`** — the value fights the document's *own* applied styles.
  The document supplies the target, so the finding says "this deviates from
  what you established elsewhere", not "this is wrong".
- **`structural`** — an objective defect: an outline that skips a level, a
  cross-reference to a bookmark that does not exist.
- **`policy`** — the value differs from a target *you* supplied. None ship
  enabled, because the library has no opinion about your house style.

`--json` emits the full finding shape — rule, kind, severity, message,
location, observed / expected — for piping into `jq` or a report.

`--no-tables` skips paragraphs inside table cells. Headers, footers,
footnotes, endnotes, and comments are not audited at all yet.

## `skill`

The library ships an [agent skill](SKILLS.md) — an LLM-facing guide — inside
the package at `docx_plus/skill/`. This command puts it where an agent will
find it.

```console
$ docx-plus skill install
installed 10 files to .claude/skills/docx-plus

$ docx-plus skill list
cli
comments
forms
layout
numbering
publishing
revisions
styles
tables

$ docx-plus skill show tables | head -3
# Tables — borders, shading, merging
```

- `skill install [--dest DIR | --user] [--force]` — copy the tree into a skills
  directory. Defaults to `./.claude/skills/`; `--user` targets
  `~/.claude/skills/`; `--dest` names any other skills root. The skill lands at
  `<root>/docx-plus`. An existing installation is left alone unless `--force`
  is given, so upgrading the library never silently clobbers local edits.
  `--user` and `--dest` are mutually exclusive.
- `skill path` — print the directory holding the packaged copy.
- `skill list` — the reference topics available to `show`.
- `skill show [TOPIC]` — print `SKILL.md`, or one reference topic, to stdout. A
  trailing `.md` on the topic is tolerated.

Because the skill ships in the wheel, `pip install docx-plus` is enough — there
is no need to clone the repository to get it.
