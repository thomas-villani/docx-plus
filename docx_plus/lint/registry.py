"""Rule registration and selection.

Rules register themselves at import time via the :func:`rule` decorator, so
adding one is a single new function in ``lint/rules/`` — no central list to
keep in sync. :func:`select_rules` implements the selection semantics the CLI
exposes.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx_plus.core import DocxPlusError
from docx_plus.lint.models import Rule

if TYPE_CHECKING:
    from collections.abc import Callable, Iterable, Sequence

    from docx_plus.lint.models import CheckFn, RuleKind, Severity
    from docx_plus.lint.profile import Profile


_REGISTRY: dict[str, Rule] = {}


class UnknownRuleError(DocxPlusError, KeyError):
    """Raised when a selector names neither a registered rule id nor a tag.

    A typo in a rule name would otherwise silently select nothing, which
    reads exactly like a clean document.
    """


def rule(
    *,
    id: str,  # noqa: A002 — "id" is the field's name in the public Finding shape
    kind: RuleKind,
    severity: Severity,
    description: str,
    tags: Iterable[str] = (),
    default_on: bool = True,
) -> Callable[[CheckFn], CheckFn]:
    """Register the decorated function as a lint rule.

    Args:
        id: Stable kebab-case identifier; part of the public surface.
        kind: ``consistency`` / ``structural`` / ``policy``.
        severity: ``error`` / ``warning`` / ``info``.
        description: One line for ``list_rules``.
        tags: Cluster names for bulk selection.
        default_on: Whether it runs when nothing is selected.

    Returns:
        The undecorated function, so rules stay directly callable in tests.

    Raises:
        ValueError: If ``id`` is already registered.
    """

    def decorate(check: CheckFn) -> CheckFn:
        if id in _REGISTRY:
            raise ValueError(f"duplicate lint rule id: {id!r}")
        _REGISTRY[id] = Rule(
            id=id,
            kind=kind,
            severity=severity,
            description=description,
            check=check,
            tags=frozenset(tags),
            default_on=default_on,
        )
        return check

    return decorate


def all_rules() -> list[Rule]:
    """Every registered rule, sorted by id."""
    _load_rules()
    return sorted(_REGISTRY.values(), key=lambda r: r.id)


def select_rules(
    select: Sequence[str] | None = None,
    exclude: Sequence[str] | None = None,
    profile: Profile | None = None,
) -> list[Rule]:
    """Resolve selectors to the rules that should run.

    Selection semantics, matching the sibling `wordlive` linter so the two
    behave the same way:

    - ``select=None`` runs every rule with ``default_on=True``, as adjusted
      by ``profile``.
    - A non-empty ``select`` runs exactly what it names, **including
      off-by-default rules and anything a profile disabled** — naming a tag
      is how a user opts into that cluster's heuristic rules, and asking
      for a rule by name is not something configuration gets to veto.
    - ``exclude`` is applied last and always wins.

    Args:
        select: Rule ids and/or tags to run.
        exclude: Rule ids and/or tags to skip.
        profile: A loaded profile whose per-rule ``enabled`` overrides the
            registered ``default_on``.

    Returns:
        The matching rules, sorted by id.

    Raises:
        UnknownRuleError: If a selector, or a rule id named by ``profile``,
            matches no registered rule or tag.
    """
    rules = all_rules()

    if profile is not None:
        _reject_unknown_ids(profile.rules, rules)

    if select:
        _reject_unknown(select, rules)
        chosen = [r for r in rules if any(r.matches(s) for s in select)]
    elif profile is not None:
        chosen = [r for r in rules if profile.enabled(r.id, default=r.default_on)]
    else:
        chosen = [r for r in rules if r.default_on]

    if exclude:
        _reject_unknown(exclude, rules)
        chosen = [r for r in chosen if not any(r.matches(s) for s in exclude)]

    return chosen


def _reject_unknown_ids(named: Iterable[str], rules: Sequence[Rule]) -> None:
    """Reject rule ids a profile names that do not exist.

    Tags are not accepted here, unlike in a selector: a profile sets
    per-rule severities and options, and "apply this severity to whatever
    happens to carry the tag today" is not a stable thing to check into a
    repository.
    """
    known = {r.id for r in rules}
    unknown = [name for name in named if name not in known]
    if unknown:
        raise UnknownRuleError(f"profile names unknown rule(s): {', '.join(sorted(unknown))}")


def _reject_unknown(selectors: Sequence[str], rules: Sequence[Rule]) -> None:
    known = {r.id for r in rules} | {tag for r in rules for tag in r.tags}
    unknown = [s for s in selectors if s not in known]
    if unknown:
        raise UnknownRuleError(
            f"unknown rule id or tag: {', '.join(sorted(unknown))}. "
            f"Known tags: {', '.join(sorted({t for r in rules for t in r.tags}))}"
        )


def _load_rules() -> None:
    """Import the rule modules so their decorators run.

    Deferred rather than imported at package init to keep the import cycle
    simple: rules import the models and the registry, so the registry cannot
    import the rules at module scope.
    """
    from docx_plus.lint import rules as _rules  # noqa: F401


__all__ = [
    "UnknownRuleError",
    "all_rules",
    "rule",
    "select_rules",
]
