# `docx_plus.lint`

Audit a document for formatting defects. See the
[CLI page](../cli.md#lint) for the `docx-plus lint` command over the same
engine.

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

::: docx_plus.lint.engine
    options:
      members:
        - lint

::: docx_plus.lint.models
    options:
      members:
        - Finding
        - Issue
        - Location
        - Rule
        - LintContext
        - RuleKind
        - Severity

::: docx_plus.lint.registry
    options:
      members:
        - rule
        - all_rules
        - select_rules
        - UnknownRuleError
