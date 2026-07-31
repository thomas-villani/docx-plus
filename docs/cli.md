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
  plan      show the edits that would repair a document
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

`lint` and `plan` overload `1` deliberately: a successful run that *found
something* also exits `1`, so both work as CI gates. A genuine failure is
still distinguishable — it prints an `error:` line to stderr and no findings.

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
    bold      : True   <- paragraphStyle: Title
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

The clusters are `typography`, `whitespace`, `structure`, `headings`,
`lists`, `formatting`, `fonts`, `styles`, `references`, `fields`,
`captions`, and `language`. Rules carry more than one, so `--rule styles`
picks up the style-definition rules *and* the direct-formatting ones that
judge against a style.

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
20 rules, 16 on by default.
```

The four that ship **off** are the ones whose thresholds are a judgement
rather than a fact — `mixed-run-formatting`, `stray-empty-paragraph`,
`font-outliers`, and `unused-styles`. Enable them by id or by tag when you
want them.

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
location, observed / expected, `fixable`, `adds_content` — for piping into
`jq` or a report.

`--no-tables` skips paragraphs inside table cells. Headers, footers,
footnotes, endnotes, and comments are not audited at all yet.

**Profiles.** `--profile PATH` applies a
[lint profile](reference/lint.md#profiles) — per-rule enable / disable and
severity overrides. Without it, `docx-plus-lint.json` is looked for beside
the document and upwards, so a repository can check its conventions in;
`--no-profile` ignores both. `lint` and `plan` take the same options.

## `plan`

Show what repairing a document would change, wrapping
[`plan_fixes`](reference/lint.md#fixes-and-the-plan). Also read-only —
**this release applies nothing.** The command exists to make the repair
inspectable before anything can perform it.

```console
$ docx-plus plan report.docx
3 edit(s), in the order they would be applied:

1. [check] paragraph 0  double-space
     Collapse 1 run of spaces to a single space.
     - replace-paragraph-text paragraph_index=0 spans=[9-11->' ']
2. [safe ] paragraph 4, run 0  redundant-direct-formatting
     Delete the run's direct size; the style supplies the same value.
     - clear-run-properties paragraph_index=4 run_index=0 properties=font_size
3. [check] paragraph 5  style-drift
     Clear the paragraph's direct space after so Normal applies.
     - clear-paragraph-properties paragraph_index=5 properties=spacing_after

2 finding(s) with no known repair:
     1  heading-level-skip
     1  manual-list

3 to apply, 0 withheld, 0 dropped, 2 unfixable.
```

The four possible sections are the four things that can happen to a finding:
it becomes an edit, it is withheld for changing content, it loses a
collision with another edit, or nobody knows how to fix it. Every finding
lands in exactly one, so the report accounts for the whole audit rather than
listing only the good news.

The bracketed mark is the fix's **safety class**: `safe` means the document
renders identically afterwards and only the XML gets tidier, `check` means
the rendering or text changes deliberately, and `DROP` means something is
removed that cannot be reconstructed.

**The content gate.** An edit that deletes a paragraph or a style definition
changes what the document *contains*, not how it looks, so it is withheld
and reported separately:

```console
$ docx-plus plan report.docx --rule stray-empty-paragraph
No edits.

1 withheld - these change what the document contains:
  paragraph 6  stray-empty-paragraph
    Delete 1 empty paragraph, keeping one.
  (re-run with --allow-content to include them)

0 to apply, 1 withheld, 0 dropped, 0 unfixable.
```

`--allow-content` includes them. `--json` emits the whole plan
— every operation with its arguments — so it can be stored, reviewed, or
diffed between runs. `--rule`, `--exclude`, `--no-tables`, `--profile`, and
`--no-profile` work exactly as they do for `lint`.

Exit code `1` when the plan holds any edit, applied or withheld, so `plan`
gates a pipeline on "is there anything a repair pass would do". Findings
nobody can repair do not fail that gate.

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
