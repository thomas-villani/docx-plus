"""Document linting — audit a ``.docx`` for formatting defects.

A **composing layer**, not a capability module. Like ``cli/``, it sits
above the capability packages and reads across them; it adds no OOXML
knowledge of its own, and every judgement it makes is built on
``styles/``'s cascade resolver.

The rules divide into three **kinds**, which is what keeps an opinionated
feature inside a lean library:

- ``consistency`` — a value fights the document's own applied styles. The
  document supplies the target, so no configuration is needed and no
  opinion is imposed: the rule only reports that something deviates from
  what was established elsewhere.
- ``structural`` — an objective defect, true regardless of house style.
- ``policy`` — a value differs from a target the *user* supplied. Inert
  without one.

So `docx_plus` reports that forty paragraphs resolve identically under
three different style ids; it does not assert that this is wrong. What to
do about it stays the author's call.

Nothing here writes. :func:`lint` reports, and :func:`plan_fixes` turns
findings into an ordered, serializable description of what a repair pass
*would* change — deliberately stopping short of applying it, so the fix
model is designed and reviewed while no code path can corrupt a document.

Example:
    >>> from docx import Document
    >>> from docx_plus.lint import lint, plan_fixes
    >>> doc = Document()
    >>> _ = doc.add_paragraph("Spaced  out .")
    >>> for finding in lint(doc):
    ...     print(finding.rule)
    double-space
    space-before-punctuation
    >>> plan = plan_fixes(lint(doc))
    >>> print(len(plan.fixes), "edits,", len(plan.conflicts), "conflicts")
    2 edits, 0 conflicts
"""

from docx_plus.lint.engine import lint
from docx_plus.lint.models import (
    Finding,
    Fix,
    FixOp,
    FixOperation,
    FixSafety,
    Issue,
    LintContext,
    Location,
    Rule,
    RuleKind,
    Severity,
)
from docx_plus.lint.plan import (
    FixConflict,
    FixPlan,
    InvalidFixError,
    PlannedFix,
    plan_fixes,
)
from docx_plus.lint.profile import (
    DEFAULT_PROFILE_NAME,
    InvalidProfileError,
    Profile,
    RuleSettings,
)
from docx_plus.lint.registry import UnknownRuleError, all_rules, rule, select_rules

__all__ = [
    "DEFAULT_PROFILE_NAME",
    "Finding",
    "Fix",
    "FixConflict",
    "FixOp",
    "FixOperation",
    "FixPlan",
    "FixSafety",
    "InvalidFixError",
    "InvalidProfileError",
    "Issue",
    "LintContext",
    "Location",
    "PlannedFix",
    "Profile",
    "Rule",
    "RuleKind",
    "RuleSettings",
    "Severity",
    "UnknownRuleError",
    "all_rules",
    "lint",
    "plan_fixes",
    "rule",
    "select_rules",
]
