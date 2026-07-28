"""The lint entry point: sweep the document once, run the selected rules."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx_plus.lint.models import Finding, LintContext
from docx_plus.lint.registry import select_rules
from docx_plus.styles import iter_resolved_paragraphs

if TYPE_CHECKING:
    from collections.abc import Sequence

    from docx.document import Document

    from docx_plus.lint.models import Rule


def lint(
    doc: Document,
    *,
    select: Sequence[str] | None = None,
    exclude: Sequence[str] | None = None,
    include_tables: bool = True,
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

    Returns:
        Findings sorted by severity, then document order.

    Raises:
        UnknownRuleError: If a selector matches no rule id or tag.
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
    rules = select_rules(select, exclude)
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
        findings.extend(_run(rule, context))

    return sorted(findings, key=lambda f: f.sort_key)


def _run(rule: Rule, context: LintContext) -> list[Finding]:
    """Run one rule, promoting each :class:`Issue` it yields to a Finding.

    A rule body supplies only what it knows — the message, where, and what
    it saw. The id, kind, and severity come from the registration, so a rule
    cannot drift from what ``list_rules`` advertises about it.
    """
    return [
        Finding(
            rule=rule.id,
            kind=rule.kind,
            severity=issue.severity or rule.severity,
            message=issue.message,
            location=issue.location,
            observed=issue.observed,
            expected=issue.expected,
            fixable=issue.fixable,
            adds_content=issue.adds_content,
        )
        for issue in rule.check(context)
    ]


__all__ = ["lint"]
