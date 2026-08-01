# Linting

`docx_plus.lint` has two entry points, both **pure reads**: `lint(doc)`
reports formatting defects, and `plan_fixes(findings)` turns those findings
into an ordered, serializable description of what a repair pass *would*
change. Nothing in this module writes to a document.

The [cascade resolver](styles.md) can already say *why* a paragraph looks
the way it does. Lint is that answer applied at document scale — direct
formatting fighting the styles, outlines that skip a level, hand-typed list
markers, whitespace used as layout.

**Use it when** you are handed someone else's `.docx` and need to know
what's wrong before editing, or when a CI step should fail on a document
that has drifted. Don't reach for it on a document you just built — you
know what you put in it.

## The one-liner

```python
from docx import Document
from docx_plus.lint import lint

doc = Document("report.docx")
for finding in lint(doc):
    print(f"{finding.severity:8} {finding.rule:28} {finding.location.describe()}")
    print(f"         {finding.message}")
```

`lint(doc, *, select=None, exclude=None, include_tables=True,
profile=None)` returns a `list[Finding]` sorted by severity, then document
order.

- `select` — rule ids and/or tags to run. `None` runs the default-on set.
  Naming a **tag** also switches on that cluster's off-by-default rules.
- `exclude` — ids and/or tags to skip; applied last, and always wins.
- `include_tables` — sweep paragraphs inside table cells (default `True`).
- `profile` — a `Profile`, or a path / dict / JSON file `Profile.load`
  accepts. See [Profiles](#profiles).

An unknown selector raises `UnknownRuleError`, not a silent no-op.

!!! note "Body only, and the index counts table paragraphs"
    Headers, footers, footnotes, endnotes, and comments are not swept — a
    clean lint says nothing about a document's header.

    `paragraph_index` counts table-cell paragraphs, because the sweep walks
    them. It is *not* an index into `doc.paragraphs`, which drops them.
    (And the CLI's `inspect` numbering is 1-based; everything here is
    0-based.)

## What a `Finding` carries

```python
finding.rule          # "style-drift" — the stable rule id
finding.kind          # "consistency" | "structural" | "policy"
finding.severity      # "error" | "warning" | "info"
finding.message       # a sentence, in the document's terms
finding.location      # Location(paragraph_index, run_index, style_id, excerpt)
finding.observed      # the value found, rendered for display
finding.expected      # what the rule expected, where that is meaningful
finding.fix           # a Fix, or None
finding.fixable       # property — exactly `finding.fix is not None`
finding.adds_content  # would repairing this change what the document *says*?
```

`Location` fields are all optional: a rule about a *style definition* — an
unused style, two styles that resolve identically — has no paragraph to
point at and reports `style_id` alone. `location.describe()` renders
whichever it has.

## Rule kinds — what keeps this honest about opinions

| Kind | Needs config? | The judgement |
|---|---|---|
| `consistency` | no | A value fights the document's **own** applied styles — the document supplies the target |
| `structural` | no | An objective defect, true regardless of house style |
| `policy` | yes | A value differs from a target **you** supplied |

No `policy` rule is ever on by default, so `docx_plus` never asserts a
house style of its own. It will report that forty paragraphs resolve
identically under three style ids; whether that is a problem stays your
call.

## The rule catalogue

Twenty rules, sixteen on by default. Get the live list with `docx-plus lint
--list-rules`, or `from docx_plus.lint import all_rules`.

| Rule | Kind | Severity | Default | Fix? |
|---|---|---|---|---|
| `broken-cross-reference` | structural | error | on | — |
| `caption-manual-numbering` | structural | warning | on | — |
| `direct-numbering-override` | consistency | warning | on | review |
| `double-space` | consistency | info | on | review |
| `duplicate-styles` | consistency | info | on | — |
| `empty-heading` | structural | warning | on | — |
| `font-outliers` | consistency | info | **off** | — |
| `heading-level-skip` | structural | warning | on | — |
| `indent-by-whitespace` | structural | warning | on | — |
| `list-numbering-continuity` | structural | warning | on | — |
| `manual-heading-formatting` | structural | warning | on | — |
| `manual-list` | structural | warning | on | — |
| `mixed-language` | consistency | info | on | review |
| `mixed-run-formatting` | consistency | info | **off** | — |
| `redundant-direct-formatting` | consistency | info | on | safe |
| `space-before-punctuation` | consistency | info | on | review |
| `stray-empty-paragraph` | structural | info | **off** | destructive |
| `style-drift` | consistency | warning | on | review |
| `trailing-whitespace` | structural | info | on | review |
| `unused-styles` | structural | info | **off** | destructive |

Tags for bulk selection: `captions`, `fields`, `fonts`, `formatting`,
`headings`, `language`, `lists`, `references`, `structure`, `styles`,
`typography`, `whitespace`.

```python
lint(doc, select=["typography"])     # the cluster, incl. its off-by-default rules
lint(doc, exclude=["double-space"])  # everything else
lint(doc, select=["styles"], exclude=["duplicate-styles"])
```

## Fixes and the plan

A rule that knows how to repair what it found attaches a `Fix`. There is no
separate "fixable" flag to keep in step — a finding is fixable exactly when
it carries one.

```python
from docx_plus.lint import lint, plan_fixes

plan = plan_fixes(lint(doc))
for planned in plan.fixes:
    print(planned.rule, planned.safety, planned.fix.summary)
```

A `Fix` is a `summary`, a `safety`, and a tuple of **operations** from a
closed vocabulary — not a callable, because a plan has to survive being
written to a file, reviewed, and applied by a different process than the
one that built it.

`FixOp` is one of `clear-run-properties`, `clear-paragraph-properties`,
`clear-paragraph-numbering`, `set-run-language`, `replace-paragraph-text`,
`delete-paragraph`, `delete-style`. Each carries JSON `args` — property
names are `ResolvedFormatting` field names, so a plan reads in the same
terms as the report that produced it.

`FixSafety` says how much trust applying it asks for:

- **`safe`** — the document renders identically afterwards; only the XML
  gets tidier. This is what `redundant-direct-formatting` produces, and it
  is provable rather than asserted.
- **`review`** — the rendering or the text changes, deliberately. The old
  value is in the finding's `observed`, so it is reversible by hand.
- **`destructive`** — something is removed that the document cannot
  reconstruct.

`plan_fixes(findings, *, allow_content=False)` decides the three things no
individual rule can, because each is a property of the *set* of findings:

- **Order.** Deletions last and back to front. Every operation names a
  position in the document *as swept*, so a deletion partway down
  invalidates every index below it.
- **The content gate.** A fix that removes a paragraph or a style
  definition changes what the document *contains*. Those are withheld
  unless you pass `allow_content=True`, and land in `plan.deferred` so they
  stay visible.
- **Conflicts.** Two rules can claim the same run property or overlapping
  spans of the same text. Claims are per property and per character span
  rather than per paragraph, so a paragraph with several unrelated defects
  is still fixable; the earlier fix wins and the loser is named in
  `plan.conflicts`.

`FixPlan` has `fixes`, `deferred`, `conflicts`, and `unfixable`. Every
finding lands in exactly one of them, so a plan accounts for the whole
audit rather than listing only the good news. `plan.operations` flattens
the kept fixes into applying order, and `plan.to_dict()` serialises the
lot.

!!! warning "Nothing applies a plan yet"
    Reading `plan.operations` and carrying the edits out is the caller's
    job in this release. There is no `docx-plus regularize`. Applying a
    plan is v0.7.

### Eleven rules are report-only, deliberately

A skipped outline level can be repaired by promoting this heading or
demoting the one above it, and those produce different documents. Two
styles that resolve identically give no reason to prefer either as the
survivor. Typed indentation needs a number the document does not contain.
Each rule's docstring says which case it is.

A plan that guessed would be the library asserting a house style — exactly
what the rule kinds exist to prevent.

## Profiles

The one place a house opinion may live. A profile enables and disables
rules and overrides severities.

```json
{
  "rules": {
    "double-space": {"enabled": false},
    "unused-styles": {"enabled": true},
    "style-drift": {"severity": "error"}
  }
}
```

```python
from docx_plus.lint import Profile, lint

lint(doc, profile="house-style.json")                              # a path
lint(doc, profile={"rules": {"double-space": {"enabled": False}}}) # inline
lint(doc, profile=Profile.discover("report.docx"))
```

Or check `docx-plus-lint.json` into the repository and both CLI commands
find it beside the document, or in any directory above it.

Precedence: a profile adjusts the defaults, an explicit `select` overrides
both, and `exclude` is applied last. Naming a rule with `--rule` therefore
overrides a profile that disabled it — configuration never gets to veto a
direct question about one document.

A profile may **not** configure a tag ("apply this severity to whatever
carries the tag today" is not a stable thing to check in), and one naming a
rule that does not exist raises `InvalidProfileError` / `UnknownRuleError`
on load rather than silently doing nothing.

## Writing your own rule

Rules register themselves at import, so a new one is a single function:

```python
from collections.abc import Iterator

from docx_plus.lint import Issue, LintContext, Location, rule


@rule(
    id="all-caps-heading",
    kind="consistency",
    severity="info",
    description="A heading typed in capitals rather than styled.",
    tags={"headings"},
    default_on=False,
)
def all_caps_heading(ctx: LintContext) -> Iterator[Issue]:
    for resolved in ctx.paragraphs:
        if resolved.formatting.outline_level is None:
            continue
        if resolved.text.isupper() and len(resolved.text) > 3:
            yield Issue(
                message="Heading is typed in capitals; use the style's caps property.",
                location=Location(
                    paragraph_index=resolved.index,
                    excerpt=ctx.excerpt(resolved.index),
                ),
                observed=resolved.text,
            )
```

A rule yields `Issue` — only what the rule itself knows. The engine
promotes each to a `Finding` by stamping on the id, kind, and severity from
the registration, so a rule cannot advertise one severity in `--list-rules`
and emit another.

`LintContext` gives you `ctx.doc`, `ctx.paragraphs` (every swept paragraph
in document order, each with its full resolve, provenance, and `baseline`),
plus `ctx.resolve(target, stop_below=...)` for a slice of the cascade the
sweep did not precompute, and `ctx.excerpt(paragraph_index)` for a
report-safe text slice.

Rules receive the **whole swept document**, not one paragraph at a time,
because the interesting rules are comparative — "this font is an outlier",
"these two styles resolve identically", "the outline skips a level". None
of those is decidable from a single paragraph.

## From the shell

```bash
docx-plus lint report.docx                    # findings, human-readable
docx-plus lint report.docx --json             # structured
docx-plus lint --list-rules                   # the catalogue; no document needed
docx-plus plan report.docx                    # the repair, described
docx-plus plan report.docx --allow-content    # include deletions
```

`lint` exits `1` when it found anything, so it drops into a CI step
directly. See the [CLI page](../cli.md) for every flag.

## Errors

- `UnknownRuleError` (`DocxPlusError`, `KeyError`) — a selector, or a
  profile, named a rule id or tag that does not exist.
- `InvalidProfileError` (`DocxPlusError`, `ValueError`) — the profile is
  unreadable or malformed.
- `StyleCascadeError` — a `basedOn` chain has a cycle or exceeds Word's
  depth limit. It comes from the resolver underneath, not from lint itself.

## See also

- [How the linter is designed](../concepts/lint.md) — rule kinds, the fix
  vocabulary, what only the planner can decide
- [Styles and the cascade](styles.md) — `stop_below` and the sweep the
  linter is built on
- Reference: [`lint`](../reference/lint.md)
- Example: `lint_document.py` — builds a document with one instance of
  several defects, then prints both the findings and the plan
