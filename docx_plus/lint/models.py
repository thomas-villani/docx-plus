"""The linter's vocabulary: ``Finding``, ``Location``, ``Rule``, ``LintContext``.

The shape here is deliberately close to the sibling `wordlive` linter's
(``../wordlive/spec-linter.md`` §4), so a document audited by either tool
reads the same way. The one structural difference is :class:`Location`:
wordlive addresses a live document by anchor id (``para:7``), which only
means anything to a running Word instance, so findings here carry an index
into the sweep plus an excerpt.
"""

from __future__ import annotations

from collections.abc import Callable, Iterator, Mapping
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, Literal

from docx_plus.styles import resolve_effective_formatting

if TYPE_CHECKING:
    from docx.document import Document
    from docx.table import _Cell
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run

    from docx_plus.styles import ResolvedFormatting, ResolvedParagraph
    from docx_plus.styles.inspect import Layer


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


_SEVERITY_RANK: dict[Severity, int] = {"error": 0, "warning": 1, "info": 2}

_ELLIPSIS_ROOM = 3
"""Characters an ``"..."`` suffix costs. Below this a `limit` cannot be
honoured at all, so truncation is skipped rather than made longer than the
input."""


def _to_ascii(text: str) -> str:
    r"""Render ``text`` using ASCII only, escaping anything outside it.

    ``"Café"`` becomes ``"Caf\xe9"``. Uses Python's own escape spelling so
    the result is unambiguous and searchable, and never silently drops a
    character — a report that quietly deleted the thing it was describing
    would be worse than one that spells it awkwardly.
    """
    if text.isascii():
        return text
    return text.encode("ascii", "backslashreplace").decode("ascii")


def render_for_report(text: str, limit: int = 60) -> str:
    r"""Make arbitrary document text safe to print in a report line.

    Collapses line breaks, makes tabs visible as ``\t``, escapes everything
    outside ASCII, then clips to ``limit`` printed characters. Every rule
    putting document-derived text into a :class:`Location` should go
    through this — :meth:`LintContext.excerpt` does it for paragraph text,
    and this is the same treatment for anything else, such as a field
    instruction.

    Internal spacing is deliberately **preserved**: several rules are about
    whitespace, and an excerpt that tidied it would hide the very thing
    being reported.
    """
    text = text.replace("\r", " ").replace("\n", " ").replace("\t", "\\t")
    text = _to_ascii(text)
    if limit < _ELLIPSIS_ROOM or len(text) <= limit:
        return text
    return text[: limit - _ELLIPSIS_ROOM] + "..."


FixSafety = Literal["safe", "review", "destructive"]
"""How much trust applying a fix asks for.

Orthogonal to :attr:`Finding.adds_content`, which is about *what* changes
(content or only formatting); this is about *how recoverable* the change is.

- ``safe`` — the document renders identically afterwards. Only the XML gets
  tidier: a property is deleted from a run and the same value arrives from
  the style instead. This is the class ``redundant-direct-formatting``
  produces, and it is provable rather than asserted — the rule found the
  property precisely by comparing against the value that would surface
  without it.
- ``review`` — the rendering or the text changes, deliberately. The old
  value is in the finding's ``observed``, so the change is reversible by
  hand.
- ``destructive`` — something is removed that the document cannot
  reconstruct: a style definition and everything it declared, a paragraph
  and its formatting.
"""


FixOp = Literal[
    "clear-run-properties",
    "clear-paragraph-properties",
    "clear-paragraph-numbering",
    "set-run-language",
    "replace-paragraph-text",
    "delete-paragraph",
    "delete-style",
]
"""The closed vocabulary a fix is expressed in.

Deliberately a **fixed set of named operations** rather than arbitrary
callables. A plan has to survive being written to JSON, read by a human,
and applied by a different process than the one that built it, and none of
that works if an edit is a Python object holding a bound method.

Each op and its ``args``:

``clear-run-properties``
    ``{"paragraph_index": int, "run_index": int, "properties": [str, ...]}``
    — delete the named direct properties from the run's ``w:rPr``.
``clear-paragraph-properties``
    ``{"paragraph_index": int, "properties": [str, ...]}`` — the same for a
    paragraph's ``w:pPr``.
``clear-paragraph-numbering``
    ``{"paragraph_index": int}`` — delete the paragraph's direct
    ``w:numPr``, so the list its style supplies applies again.
``set-run-language``
    ``{"paragraph_index": int, "run_index": int, "lang": str}``.
``replace-paragraph-text``
    ``{"paragraph_index": int, "spans": [{"start": int, "end": int,
    "replacement": str}, ...]}`` — half-open character spans into the
    paragraph's text, all measured against the **original** text, so the
    order they are applied in does not matter and two rules' spans can be
    checked for overlap.
``delete-paragraph``
    ``{"paragraph_index": int}``.
``delete-style``
    ``{"style_id": str}``.

Property names are :class:`~docx_plus.styles.ResolvedFormatting` field
names — the vocabulary the finding already reported in, so a plan reads in
the same terms as the report that produced it.
"""


@dataclass(frozen=True)
class FixOperation:
    """One named edit, with JSON-serializable arguments.

    Attributes:
        op: Which operation, from the :data:`FixOp` vocabulary.
        args: Its arguments. Restricted to JSON types so a plan round-trips
            through a file.
    """

    op: FixOp
    args: Mapping[str, Any]

    def to_dict(self) -> dict[str, Any]:
        """The serializable record for this operation."""
        return {"op": self.op, "args": dict(self.args)}


@dataclass(frozen=True)
class Fix:
    """What a rule would do about the thing it found.

    A fix is **described, never executed**, in this release: ``lint`` and
    :func:`~docx_plus.lint.plan_fixes` are both pure reads. Describing it
    first is the point — the fix model gets designed and reviewed while
    nothing can yet corrupt a document.

    Attributes:
        summary: One line, in the document's terms: what would change.
        safety: See :data:`FixSafety`.
        operations: The edits, **in the order they must be applied**. A rule
            supplying more than one is asserting that order matters — a run
            of empty paragraphs is deleted back to front so the earlier
            indices stay valid.
    """

    summary: str
    safety: FixSafety
    operations: tuple[FixOperation, ...]

    def to_dict(self) -> dict[str, Any]:
        """The serializable record for this fix."""
        return {
            "summary": self.summary,
            "safety": self.safety,
            "operations": [op.to_dict() for op in self.operations],
        }


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
        fix: What would repair it, or ``None`` where no unambiguous repair
            exists. There is no separate "fixable" flag to keep in step:
            a finding is fixable exactly when it carries a fix.
        adds_content: Whether the eventual fix would insert or delete
            content rather than only change formatting. Set independently of
            :attr:`fix`, so a rule can say "repairing this would change what
            the document says" while leaving the repair itself unmodelled.
    """

    message: str
    location: Location = field(default_factory=Location)
    observed: str | None = None
    expected: str | None = None
    severity: Severity | None = None
    fix: Fix | None = None
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
        fix: What would repair it, or ``None`` where no unambiguous repair
            exists — which is most of the outline and reference rules, since
            "the outline skips a level" does not say whether to promote the
            heading or demote the one above it.
        adds_content: Whether the fix would insert or delete content rather
            than only change formatting. :func:`~docx_plus.lint.plan_fixes`
            withholds those unless a caller opts in.
    """

    rule: str
    kind: RuleKind
    severity: Severity
    message: str
    location: Location = field(default_factory=Location)
    observed: str | None = None
    expected: str | None = None
    fix: Fix | None = None
    adds_content: bool = False

    @property
    def fixable(self) -> bool:
        """Whether a repair is known — exactly whether :attr:`fix` is set."""
        return self.fix is not None

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

    def resolve(
        self,
        target: Paragraph | Run | _Cell,
        *,
        stop_below: Layer | None = None,
    ) -> ResolvedFormatting:
        """Resolve one target the sweep did not precompute.

        The sweep already carries each paragraph's and run's full resolve
        plus its ``baseline`` (the same target without its own direct
        layer), which is what nearly every rule needs. This covers the rest:
        a rule wanting some *other* slice of the cascade — the numbering a
        style would supply if the paragraph did not override it, say.

        Deliberately not cache-shared with the sweep, so it costs a full
        cascade walk per call. Call it for the subset a rule genuinely
        needs, never for every paragraph.

        Args:
            target: A paragraph, run, or cell in :attr:`doc`.
            stop_below: Stop the walk below this
                :data:`~docx_plus.styles.inspect.Layer`.

        Returns:
            The resolved formatting for ``target``.
        """
        return resolve_effective_formatting(target, stop_below=stop_below)

    def excerpt(self, paragraph_index: int, limit: int = 60) -> str:
        r"""A one-line slice of a paragraph's text, for a report.

        Internal spacing is **preserved**, because several rules are about
        whitespace and an excerpt that tidied it would hide the very thing
        being reported — a `double-space` finding whose excerpt shows single
        spaces reads like a false positive. Tabs render as ``\t`` so they are
        visible at all, and line breaks collapse to a space so one finding
        stays one line.

        Output is **ASCII-only, and enforced rather than hoped for**. This
        reaches a Windows console via ``docx-plus lint``, where Python
        encodes stdout as cp1252 whenever it is redirected to a file or a
        pipe — so a document containing CJK text or an emoji used to end
        the command in an unhandled ``UnicodeEncodeError`` rather than a
        report. Anything outside ASCII becomes a ``\x``/``\u`` escape,
        which keeps the character visible and greppable instead of
        dropping it.

        Truncation happens **after** escaping, so ``limit`` bounds the
        printed width rather than the source length. See
        :func:`render_for_report`, which is the same treatment for text
        that does not come from a paragraph.
        """
        return render_for_report(self.paragraphs[paragraph_index].text, limit)


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
    "Fix",
    "FixOp",
    "FixOperation",
    "FixSafety",
    "Issue",
    "LintContext",
    "Location",
    "Rule",
    "RuleKind",
    "Severity",
]
