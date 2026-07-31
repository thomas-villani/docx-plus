"""The lint entry point: sweep the document once, run the selected rules."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx_plus.lint.models import Finding, LintContext
from docx_plus.lint.profile import Profile
from docx_plus.lint.registry import select_rules
from docx_plus.styles import iter_resolved_paragraphs

if TYPE_CHECKING:
    from collections.abc import Mapping, Sequence
    from pathlib import Path
    from typing import Any

    from docx.document import Document

    from docx_plus.lint.models import Rule


def lint(
    doc: Document,
    *,
    select: Sequence[str] | None = None,
    exclude: Sequence[str] | None = None,
    include_tables: bool = True,
    profile: Profile | str | Path | Mapping[str, Any] | None = None,
) -> list[Finding]:
    """Audit ``doc`` and return what the selected rules noticed.

    Pure read: nothing here mutates the document. The cascade is resolved
    once for the whole document and shared across every rule, so the cost is
    one sweep regardless of how many rules run.

    Args:
        doc: The python-docx :class:`~docx.document.Document` to audit.
        select: Rule ids and/or tags to run. ``None`` runs the
            default-on set. Naming a tag also enables that cluster's
            off-by-default rules.
        exclude: Rule ids and/or tags to skip; applied last.
        include_tables: Whether to sweep paragraphs inside table cells.
        profile: A :class:`~docx_plus.lint.Profile`, or anything
            :meth:`~docx_plus.lint.Profile.load` accepts. Supplies a team's
            enable/disable and severity overrides. ``select`` and
            ``exclude`` are applied *after* it, so naming a rule explicitly
            always wins over what a profile said about it.

    Returns:
        Findings sorted by severity, then document order.

    Raises:
        UnknownRuleError: If a selector matches no rule id or tag.
        InvalidProfileError: If ``profile`` is malformed.
        StyleCascadeError: If a ``basedOn`` chain has a cycle or exceeds
            Word's depth limit.

    Note:
        Only the main document body is audited. Headers, footers,
        footnotes, endnotes, and comments are not swept — see
        :func:`~docx_plus.styles.iter_resolved_paragraphs`.

    Example:
        >>> from docx import Document
        >>> from docx_plus.lint import lint
        >>> doc = Document()
        >>> _ = doc.add_paragraph("Two  spaces here.")
        >>> for finding in lint(doc):
        ...     print(finding.rule, "-", finding.message)
        double-space - Two or more consecutive spaces between words.
    """
    loaded = profile if isinstance(profile, Profile) else Profile.load(profile)
    rules = select_rules(select, exclude, loaded)
    # Provenance and baselines are always on. Neither is an optional extra
    # here: the consistency rules are built on knowing *which* cascade layer
    # set a value and what the value would have been without it, which is
    # the whole advantage of resolving OOXML rather than asking Word for an
    # effective number. One sweep serves every rule, so the cost is paid
    # once however many rules run.
    context = LintContext(
        doc=doc,
        paragraphs=list(
            iter_resolved_paragraphs(
                doc,
                include_provenance=True,
                include_baseline=True,
                include_tables=include_tables,
            )
        ),
    )

    findings: list[Finding] = []
    for rule in rules:
        findings.extend(_run(rule, context, loaded))

    return sorted(findings, key=lambda f: f.sort_key)


def _run(rule: Rule, context: LintContext, profile: Profile) -> list[Finding]:
    """Run one rule, promoting each :class:`Issue` it yields to a Finding.

    A rule body supplies only what it knows — the message, where, and what
    it saw. The id, kind, and severity come from the registration, so a rule
    cannot drift from what ``list_rules`` advertises about it.

    An :class:`Issue` that names its own severity keeps it. A rule
    downgrading one finding is saying something about *that* finding — the
    numbering opt-out sentinel is a legitimate override, unlike every other
    one the rule reports — and a profile's blanket "treat this rule as an
    error" is not an answer to that.
    """
    default = profile.severity(rule.id, default=rule.severity)
    return [
        Finding(
            rule=rule.id,
            kind=rule.kind,
            severity=issue.severity or default,
            message=issue.message,
            location=issue.location,
            observed=issue.observed,
            expected=issue.expected,
            fix=issue.fix,
            adds_content=issue.adds_content,
        )
        for issue in rule.check(context)
    ]


__all__ = ["lint"]
