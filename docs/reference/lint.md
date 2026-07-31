# `docx_plus.lint`

Audit a document for formatting defects, and describe what repairing them
would change. See the [CLI page](../cli.md#lint) for the `docx-plus lint`
and `docx-plus plan` commands over the same engine.

**Nothing here writes.** `lint` reports; `plan_fixes` turns findings into an
ordered, serializable description of the repair and stops there. Designing
the fix model at a point where no code path can apply it is the whole reason
the two halves shipped separately.

A **composing layer**, not a capability module: like `cli/` it sits above the
capability packages and reads across them, adding no OOXML knowledge of its
own. Every judgement it makes is built on
[`styles/`'s cascade resolver](styles-inspect.md) and the
[document sweep](styles-sweep.md).

## Rule kinds

The distinction that keeps an opinionated feature inside a lean library:

| Kind | Needs config? | The judgement |
|---|---|---|
| `consistency` | no | a value fights the document's **own** applied styles — the document supplies the target |
| `structural` | no | an objective defect, true regardless of house style |
| `policy` | yes | a value differs from a target **you** supplied |

No `policy` rule is ever enabled by default, so `docx_plus` never asserts a
house style of its own. It reports that forty paragraphs resolve identically
under three style ids; whether that is a problem stays your call.

## Writing a rule

Rules register themselves at import, so a new one is a single function:

```python
from collections.abc import Iterator

from docx_plus.lint import Issue, LintContext, rule


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
        text = resolved.text
        if text.isupper() and len(text) > 3:
            yield Issue(
                message="Heading is typed in capitals; use the style's caps property.",
                location=Location(paragraph_index=resolved.index),
                observed=text,
            )
```

A rule yields [`Issue`][docx_plus.lint.models.Issue] — only what the rule
itself knows. The engine promotes each to a
[`Finding`][docx_plus.lint.models.Finding] by stamping on the id, kind, and
severity from the registration, so a rule cannot advertise one severity in
`--list-rules` and emit another.

Rules receive the **whole swept document**, not one paragraph at a time,
because the interesting rules are comparative: "this font is an outlier",
"these two styles resolve identically", "the outline skips a level". None of
those can be decided from a single paragraph.

## Fixes and the plan

A rule that knows how to repair what it found attaches a
[`Fix`][docx_plus.lint.models.Fix] to its `Issue`. There is no separate
"fixable" flag to keep in step: a finding is fixable exactly when it carries
one.

```python
from docx_plus.lint import lint, plan_fixes

plan = plan_fixes(lint(doc))
for planned in plan.fixes:
    print(planned.rule, planned.safety, planned.fix.summary)
```

A fix is a sequence of named **operations** from a closed vocabulary
([`FixOp`][docx_plus.lint.models.FixOp]), not a callable — a plan has to
survive being written to a file, reviewed, and handed to a different process
than the one that built it.

[`plan_fixes`][docx_plus.lint.plan.plan_fixes] then decides the three things
no individual rule can, because each is a property of the *set* of findings:

- **Order.** Deletions last and back to front. Every operation names a
  position in the document as it was swept, so a deletion partway down
  invalidates every index below it.
- **The content gate.** A fix that removes a paragraph or a style definition
  changes what the document *contains*, not how it looks. Those are withheld
  unless you pass `allow_content=True`, and reported in `plan.deferred` so
  they are visible rather than silently dropped.
- **Conflicts.** Two rules can independently claim the same run property or
  overlapping spans of the same text. Claims are per property and per
  character span rather than per paragraph, so a paragraph carrying several
  unrelated defects is still fixable; the earlier fix wins and the loser is
  named in `plan.conflicts`.

Every finding lands in exactly one of `fixes`, `deferred`, `conflicts`, or
`unfixable`, so a plan accounts for the whole audit.

### Not everything has a fix

Eleven of the twenty rules are report-only, and deliberately. A skipped
outline level can be repaired by promoting this heading or demoting the one
above it, and those produce different documents; two styles that resolve
identically give no reason to prefer either as the survivor; typed
indentation needs a number the document does not contain. Each rule's
docstring says which case it is. A plan that guessed would be the library
asserting a house style, which is exactly what the rule kinds exist to
prevent.

## Profiles

The one place a house opinion may live. A profile enables and disables
rules, overrides severities, and — once `policy` rules ship — supplies their
targets.

```json
{
  "rules": {
    "double-space": {"enabled": false},
    "style-drift":  {"severity": "error"}
  }
}
```

Pass it as `lint(doc, profile=...)`, or check `docx-plus-lint.json` into the
repository and both CLI commands will find it beside the document (or above
it). Naming a rule with `--rule` overrides a profile that disabled it:
configuration never gets to veto a direct question about one document.

A profile may not configure a **tag** — "apply this severity to whatever
carries the tag today" is not a stable thing to check in — and a profile
naming a rule that does not exist is an error on load rather than a setting
that silently does nothing.

::: docx_plus.lint.engine
    options:
      members:
        - lint

::: docx_plus.lint.plan
    options:
      members:
        - plan_fixes
        - FixPlan
        - PlannedFix
        - FixConflict

::: docx_plus.lint.models
    options:
      members:
        - Finding
        - Issue
        - Location
        - Rule
        - LintContext
        - Fix
        - FixOperation
        - FixOp
        - FixSafety
        - RuleKind
        - Severity

::: docx_plus.lint.profile
    options:
      members:
        - Profile
        - RuleSettings
        - InvalidProfileError

::: docx_plus.lint.registry
    options:
      members:
        - rule
        - all_rules
        - select_rules
        - UnknownRuleError
