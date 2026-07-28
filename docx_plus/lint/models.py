"""The linter's vocabulary: ``Finding``, ``Location``, ``Rule``, ``LintContext``.

The shape here is deliberately close to the sibling `wordlive` linter's
(``../wordlive/spec-linter.md`` §4), so a document audited by either tool
reads the same way. The one structural difference is :class:`Location`:
wordlive addresses a live document by anchor id (``para:7``), which only
means anything to a running Word instance, so findings here carry an index
into the sweep plus an excerpt.
"""

from __future__ import annotations

from collections.abc import Callable, Iterator
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Literal

if TYPE_CHECKING:
    from docx.document import Document

    from docx_plus.styles import ResolvedParagraph


Severity = Literal["error", "warning", "info"]
"""How much a finding matters. ``error`` is a defect that will misrender or
resolve wrongly; ``warning`` is a real inconsistency; ``info`` is a nudge."""


RuleKind = Literal["consistency", "structural", "policy"]
"""What *kind* of judgement a rule makes — the distinction that keeps this
layer honest about opinions.

- ``consistency`` — a value fights the document's own applied styles. Needs
  no configuration, because the document supplies the target: the rule only
  says "this deviates from what you established elsewhere".
- ``structural`` — an objective defect, true regardless of house style: an
  outline that skips a level, a `REF` to a bookmark that does not exist.
- ``policy`` — a value differs from a target the *user* supplied via a
  profile. Inert without one, so the library ships no opinion of its own.
"""


_SEVERITY_RANK: dict[str, int] = {"error": 0, "warning": 1, "info": 2}


@dataclass(frozen=True)
class Location:
    """Where a finding sits.

    Every field is optional because findings are not all positional: a rule
    about a *style definition* (an unused style, two styles that resolve
    identically) has no paragraph to point at, and reports ``style_id``
    alone.

    Attributes:
        paragraph_index: Position in the sweep's document order — the
            ``index`` of the :class:`~docx_plus.styles.ResolvedParagraph`
            the finding came from. Note this counts table-cell paragraphs,
            which ``doc.paragraphs`` omits.
        run_index: Position within that paragraph, for a run-level finding.
        style_id: The ``w:styleId`` a finding is about, when the subject is
            a style rather than a position.
        excerpt: A short slice of the paragraph's text, so a report reads
            usefully without the document open alongside it.
    """

    paragraph_index: int | None = None
    run_index: int | None = None
    style_id: str | None = None
    excerpt: str = ""

    def describe(self) -> str:
        """A short human-readable position, for report lines."""
        if self.paragraph_index is None:
            return f"style {self.style_id}" if self.style_id else "document"
        where = f"paragraph {self.paragraph_index}"
        if self.run_index is not None:
            where += f", run {self.run_index}"
        return where


@dataclass(frozen=True)
class Issue:
    """What a rule body yields — the parts only the rule knows.

    The rule's id, kind, and severity come from its registration, so a rule
    never restates them and they cannot drift from what ``list_rules``
    advertises. The engine promotes each :class:`Issue` to a
    :class:`Finding` by stamping that metadata on.

    Attributes:
        message: A sentence describing the problem, in the document's terms.
        location: Where it is.
        observed: The value found, rendered for display.
        expected: The value the rule would have expected, where that is
            meaningful.
        severity: Overrides the rule's default severity for this one
            finding, for rules whose seriousness depends on what they found.
        fixable: Whether a fix is known.
        adds_content: Whether the eventual fix would insert or delete
            content rather than only change formatting.
    """

    message: str
    location: Location = field(default_factory=Location)
    observed: str | None = None
    expected: str | None = None
    severity: Severity | None = None
    fixable: bool = False
    adds_content: bool = False


@dataclass(frozen=True)
class Finding:
    """One thing a rule noticed.

    Attributes:
        rule: The stable id of the rule that produced it.
        kind: The rule's :data:`RuleKind`.
        severity: The rule's :data:`Severity`.
        message: A sentence describing the problem, in the document's terms.
        location: Where it is.
        observed: The value found, rendered for display.
        expected: The value the rule would have expected, where that is
            meaningful. ``None`` for rules that report a shape rather than a
            mismatch.
        fixable: Whether a fix is known. Always ``False`` in v0.6 — the
            plan/apply half is v0.7 — but part of the shape now so rules do
            not have to be rewritten to gain one.
        adds_content: Whether the eventual fix would insert or delete
            content rather than only change formatting. Such fixes are
            withheld unless a caller opts in.
    """

    rule: str
    kind: RuleKind
    severity: Severity
    message: str
    location: Location = field(default_factory=Location)
    observed: str | None = None
    expected: str | None = None
    fixable: bool = False
    adds_content: bool = False

    @property
    def sort_key(self) -> tuple[int, int, int]:
        """Severity first, then document order — the report's natural order."""
        return (
            _SEVERITY_RANK[self.severity],
            self.location.paragraph_index if self.location.paragraph_index is not None else -1,
            self.location.run_index if self.location.run_index is not None else -1,
        )


@dataclass
class LintContext:
    """Everything a rule is given to work with.

    Rules receive the whole swept document rather than one paragraph at a
    time, because the interesting rules are comparative — "this font is an
    outlier", "these two styles resolve identically", "the outline skips a
    level" — and none of those can be decided from a single paragraph.

    Attributes:
        doc: The document, for rules that need a part the sweep does not
            cover (the styles element, the bookmark table).
        paragraphs: Every swept paragraph, in document order, materialised
            so rules can walk it more than once.
    """

    doc: Document
    paragraphs: list[ResolvedParagraph]

    def excerpt(self, paragraph_index: int, limit: int = 60) -> str:
        r"""A one-line slice of a paragraph's text, for a report.

        Internal spacing is **preserved**, because several rules are about
        whitespace and an excerpt that tidied it would hide the very thing
        being reported — a `double-space` finding whose excerpt shows single
        spaces reads like a false positive. Tabs render as ``\t`` so they are
        visible at all, and line breaks collapse to a space so one finding
        stays one line.

        Output is ASCII-only: this reaches a Windows console at cp1252 via
        ``docx-plus lint``.
        """
        text = self.paragraphs[paragraph_index].text
        text = text.replace("\r", " ").replace("\n", " ").replace("\t", "\\t")
        return text if len(text) <= limit else text[: limit - 3] + "..."


CheckFn = Callable[[LintContext], Iterator[Issue]]
"""What a rule's body looks like: swept document in, issues out."""


@dataclass(frozen=True)
class Rule:
    """A registered rule — its metadata and the check that implements it.

    Attributes:
        id: Stable, kebab-case, and part of the public surface: users
            select and exclude by it, so it should not change once shipped.
        kind: See :data:`RuleKind`.
        severity: Default severity of the findings it emits.
        description: One line, shown by ``list_rules``.
        check: The implementation.
        tags: Cluster names a user can select instead of naming ids
            (``typography``, ``structure``, ...). Naming a tag also
            **enables** that cluster's off-by-default rules.
        default_on: Whether the rule runs when the caller selects nothing.
            Unambiguous defects ship on; heuristic or opinion-flavoured
            rules ship off, so default output stays worth reading. Every
            ``policy`` rule is off, since it has no target without a
            profile.
    """

    id: str
    kind: RuleKind
    severity: Severity
    description: str
    check: CheckFn
    tags: frozenset[str] = frozenset()
    default_on: bool = True

    def matches(self, selector: str) -> bool:
        """True if ``selector`` names this rule, by id or by one of its tags."""
        return selector == self.id or selector in self.tags


__all__ = [
    "CheckFn",
    "Finding",
    "Issue",
    "LintContext",
    "Location",
    "Rule",
    "RuleKind",
    "Severity",
]
