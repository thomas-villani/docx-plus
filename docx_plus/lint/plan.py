"""Report to plan: turn findings into an ordered, inspectable list of edits.

**Nothing here writes.** :func:`plan_fixes` takes findings and returns a
description of what a repair pass *would* change — the ordering it would
apply the edits in, the ones it would withhold, and the pairs that cannot
both be applied. Applying a plan is a later release, and building the plan
first is deliberate: the whole fix model gets designed, serialized, and
reviewed at a point where no code path can corrupt a document.

Three things the planner owns that no individual rule can decide, because
each of them is a property of the *set* of findings rather than of any one:

- **Order.** Deletions go last and run back to front. A rule that deletes
  paragraph 12 and a rule that edits the text of paragraph 40 both address
  positions in the document as swept, and the first deletion invalidates
  every index after it. Sorting deletions to the end and applying them in
  descending order keeps every other edit's position valid.
- **The content gate.** A fix that removes a paragraph or a style
  definition changes what the document *contains*, not how it looks. Those
  are withheld unless the caller asks for them, and reported separately so
  they are visible rather than silently dropped.
- **Conflicts.** Two rules can independently target the same run, the same
  paragraph property, or overlapping spans of the same text. Each is right
  on its own and they cannot both apply.

A plan is JSON-serializable end to end (:meth:`FixPlan.to_dict`), because a
plan that cannot be written to a file, reviewed, and handed to a different
process is not much of a plan.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

from docx_plus.core import DocxPlusError

if TYPE_CHECKING:
    from collections.abc import Iterable, Sequence

    from docx_plus.lint.models import Finding, Fix, FixOperation, FixSafety


# Operations that remove something. They order last and descending, and
# they claim the whole of whatever they remove, so nothing else in the plan
# may touch it.
class InvalidFixError(DocxPlusError, ValueError):
    """A rule produced a fix the plan cannot order safely.

    Raised only for a fix that both deletes and does positional work — see
    :func:`_reject_mixed_deletion`. It is a rule-authoring error rather
    than anything about the document, so it surfaces at ``plan_fixes``
    rather than being carried in the plan.
    """


_DELETING_OPS = frozenset({"delete-paragraph", "delete-style"})


@dataclass(frozen=True)
class _Claim:
    """What one operation asserts exclusive use of.

    Two claims collide only within the same ``scope`` — the paragraph or
    the style they are about — and then only if they are about the same
    part of it. See :func:`_collides`.

    Attributes:
        scope: ``("paragraph", index)`` or ``("style", style_id)``.
        kind: ``whole`` (the scope is removed), ``run-property``,
            ``paragraph-property``, or ``text``.
        prop: The property name, for the two property kinds.
        run: The run index, for ``run-property``.
        span: The half-open character span, for ``text``.
    """

    scope: tuple[str, str]
    kind: str
    prop: str = ""
    run: int = -1
    span: tuple[int, int] = (-1, -1)


@dataclass(frozen=True)
class PlannedFix:
    """One finding's fix, placed in a plan.

    Position in :attr:`FixPlan.fixes` *is* the order — there is no separate
    sequence number to fall out of step with it.

    Attributes:
        finding: What was reported, carried whole so a plan reads without
            the report alongside it.
        fix: The repair. Non-optional here, unlike on the finding: a plan
            only ever holds findings that have one.
    """

    finding: Finding
    fix: Fix

    @property
    def rule(self) -> str:
        """The id of the rule that produced the finding."""
        return self.finding.rule

    @property
    def safety(self) -> FixSafety:
        """How much trust applying this fix asks for."""
        return self.fix.safety

    @property
    def adds_content(self) -> bool:
        """Whether this fix changes what the document contains."""
        return self.finding.adds_content

    @property
    def operations(self) -> tuple[FixOperation, ...]:
        """The edits, in the order the rule requires them applied."""
        return self.fix.operations

    @property
    def deletes(self) -> bool:
        """Whether any operation removes a paragraph or a style."""
        return any(op.op in _DELETING_OPS for op in self.fix.operations)

    @property
    def lowest_touched_index(self) -> int:
        """The largest ``paragraph_index`` any of this fix's operations names.

        What the back-to-front deletion order has to sort on. Sorting on the
        *finding's* location instead was unsound: a finding located at
        paragraph 1 whose fix deletes paragraph 20, planned alongside one
        located at paragraph 5 deleting paragraph 6, came out as
        ``[delete 6, delete 20]`` — and after 6 goes, the old 20 sits at 19.
        No shipped rule produces that shape, but ``plan_fixes`` is public
        and takes arbitrary findings.

        Falls back to the finding's own location for a fix whose operations
        name no paragraph at all, such as ``delete-style``.
        """
        indices = [
            index
            for op in self.fix.operations
            if isinstance(index := op.args.get("paragraph_index"), int)
        ]
        if indices:
            return max(indices)
        location = self.finding.location
        return location.paragraph_index if location.paragraph_index is not None else -1

    def to_dict(self) -> dict[str, Any]:
        """The serializable record for this planned fix."""
        return {
            "rule": self.rule,
            "severity": self.finding.severity,
            "message": self.finding.message,
            "where": self.finding.location.describe(),
            "location": {
                "paragraph_index": self.finding.location.paragraph_index,
                "run_index": self.finding.location.run_index,
                "style_id": self.finding.location.style_id,
                "excerpt": self.finding.location.excerpt,
            },
            "adds_content": self.adds_content,
            "fix": self.fix.to_dict(),
        }


@dataclass(frozen=True)
class FixConflict:
    """Two fixes that cannot both be applied.

    Resolution is by plan order and nothing else: the earlier fix is kept
    and the later one is dropped. That is arbitrary between two equally good
    repairs, which is exactly why the loser is reported rather than
    discarded — the caller can exclude the winning rule and re-plan if they
    wanted the other one.

    Attributes:
        kept: The fix that stays in the plan.
        dropped: The fix removed because of the collision.
        reason: What the two both claimed, in the document's terms.
    """

    kept: PlannedFix
    dropped: PlannedFix
    reason: str

    def to_dict(self) -> dict[str, Any]:
        """The serializable record for this conflict."""
        return {
            "reason": self.reason,
            "kept": {"rule": self.kept.rule, "where": self.kept.finding.location.describe()},
            "dropped": {
                "rule": self.dropped.rule,
                "where": self.dropped.finding.location.describe(),
            },
        }


@dataclass(frozen=True)
class FixPlan:
    """An ordered, inspectable description of what a repair pass would do.

    Attributes:
        fixes: What would be applied, in application order.
        deferred: Fixes withheld by the content gate. Not a failure — a
            caller who wants them passes ``allow_content=True``.
        conflicts: Pairs that could not both apply, with the loser named.
        unfixable: Findings with no known repair. Carried so a plan accounts
            for every finding it was given rather than quietly shortening
            the list.
    """

    fixes: tuple[PlannedFix, ...] = ()
    deferred: tuple[PlannedFix, ...] = ()
    conflicts: tuple[FixConflict, ...] = ()
    unfixable: tuple[Finding, ...] = ()

    @property
    def operations(self) -> tuple[FixOperation, ...]:
        """Every operation that would be applied, flattened into plan order."""
        return tuple(op for planned in self.fixes for op in planned.operations)

    def to_dict(self) -> dict[str, Any]:
        """The serializable record for the whole plan."""
        return {
            "fixes": [planned.to_dict() for planned in self.fixes],
            "deferred": [planned.to_dict() for planned in self.deferred],
            "conflicts": [conflict.to_dict() for conflict in self.conflicts],
            "unfixable": [
                {
                    "rule": finding.rule,
                    "severity": finding.severity,
                    "message": finding.message,
                    "where": finding.location.describe(),
                }
                for finding in self.unfixable
            ],
        }


def plan_fixes(
    findings: Sequence[Finding],
    *,
    allow_content: bool = False,
) -> FixPlan:
    """Order the repairs for ``findings`` and say which of them can coexist.

    A pure function of the findings: it never reads the document, so a plan
    can be built from a stored report. That is also its limit — it can only
    reason about what the findings say, which is why the fix vocabulary
    measures text spans against the original text rather than describing
    edits as transformations to replay.

    Args:
        findings: What :func:`~docx_plus.lint.lint` reported. Findings with
            no fix are carried through to :attr:`FixPlan.unfixable`.
        allow_content: Whether to include fixes that change what the
            document contains rather than only how it looks. Off by default:
            a formatting pass should not quietly delete a paragraph.

    Returns:
        The plan. Nothing is applied.

    Raises:
        InvalidFixError: If a fix both deletes a paragraph and does
            positional work elsewhere. See that error for why the plan
            cannot order such a fix safely.

    Example:
        >>> from docx import Document
        >>> from docx_plus.lint import lint, plan_fixes
        >>> doc = Document()
        >>> _ = doc.add_paragraph("Spaced  out .")
        >>> plan = plan_fixes(lint(doc))
        >>> for planned in plan.fixes:
        ...     print(planned.rule, "-", planned.fix.summary)
        double-space - Collapse 1 run of spaces to a single space.
        space-before-punctuation - Remove the whitespace before 1 punctuation mark.
    """
    planned = [
        PlannedFix(finding=finding, fix=finding.fix)
        for finding in findings
        if finding.fix is not None
    ]
    for candidate in planned:
        _reject_mixed_deletion(candidate)
    unfixable = tuple(finding for finding in findings if finding.fix is None)

    ordered = sorted(planned, key=_order_key)
    gated = [p.adds_content and not allow_content for p in ordered]
    withheld = tuple(p for p, out in zip(ordered, gated, strict=True) if out)
    candidates = [p for p, out in zip(ordered, gated, strict=True) if not out]

    kept, conflicts = _resolve_conflicts(candidates)
    return FixPlan(
        fixes=tuple(kept),
        deferred=withheld,
        conflicts=tuple(conflicts),
        unfixable=unfixable,
    )


def _reject_mixed_deletion(planned: PlannedFix) -> None:
    """Refuse a fix that deletes a paragraph *and* edits one somewhere else.

    The plan orders whole fixes, not individual operations, so such a fix
    has to sit in exactly one of the two phases and is wrong in either. Put
    it in the deletion phase and its non-deleting operation runs after
    other deletions have shifted the index it names; put it in the first
    phase and its own deletion runs before them, shifting theirs.

    A rule wanting both should emit two findings. Loud rather than silent,
    because ``rule`` is public: a third-party rule getting this wrong would
    otherwise produce a plan that quietly corrupts a document.
    """
    if not planned.deletes:
        return
    stray = [op.op for op in planned.operations if op.op not in _DELETING_OPS]
    if stray:
        raise InvalidFixError(
            f"{planned.rule}: a fix that deletes cannot also carry "
            f"{', '.join(sorted(set(stray)))} — emit two findings instead"
        )


def _order_key(planned: PlannedFix) -> tuple[int, int, int, str, str]:
    """Sort deletions last and back to front, everything else in document order.

    The two-phase split is the one piece of sequencing a plan cannot leave
    to the caller. Every operation addresses a position in the document as
    it was swept, so a deletion at paragraph 12 shifts everything below it;
    running deletions last and in descending order means no edit is ever
    applied against an index some earlier edit moved.

    The descending phase sorts on
    :attr:`PlannedFix.lowest_touched_index` — the deepest index the fix's
    *operations* name — not on where the finding was reported. The two
    differ, and only the former makes the guarantee true.
    """
    location = planned.finding.location
    paragraph = location.paragraph_index if location.paragraph_index is not None else -1
    run = location.run_index if location.run_index is not None else -1
    if planned.deletes:
        return (1, -planned.lowest_touched_index, 0, planned.rule, location.style_id or "")
    return (0, paragraph, run, planned.rule, "")


def _resolve_conflicts(
    candidates: Sequence[PlannedFix],
) -> tuple[list[PlannedFix], list[FixConflict]]:
    """Keep the first fix to claim anything; report every later one that collides."""
    kept: list[PlannedFix] = []
    claimed: list[tuple[_Claim, PlannedFix]] = []
    conflicts: list[FixConflict] = []

    for candidate in candidates:
        claims = list(_claims(candidate.operations))
        collision = next(
            (
                (claim, owner, mine)
                for claim, owner in claimed
                for mine in claims
                if _collides(claim, mine)
            ),
            None,
        )
        if collision is not None:
            claim, owner, mine = collision
            # Where one side removes the thing outright, that is the reason
            # worth printing: "paragraph 3 is removed" says why the other
            # edit cannot happen, in a way "paragraph 3, characters 0-4"
            # does not.
            reason = _describe(mine if mine.kind == "whole" else claim)
            conflicts.append(FixConflict(kept=owner, dropped=candidate, reason=reason))
            continue
        kept.append(candidate)
        claimed.extend((claim, candidate) for claim in claims)

    return kept, conflicts


def _claims(operations: Iterable[FixOperation]) -> list[_Claim]:
    """What a fix's operations assert exclusive use of.

    A claim is deliberately finer than "this paragraph": two rules clearing
    two different properties of the same run do not conflict, and saying
    they do would make the common case — a paragraph carrying several
    unrelated defects — look unfixable.
    """
    claims: list[_Claim] = []
    for operation in operations:
        args = operation.args
        if operation.op == "delete-style":
            claims.append(_Claim(scope=("style", str(args["style_id"])), kind="whole"))
            continue

        scope = ("paragraph", str(int(args["paragraph_index"])))
        if operation.op == "delete-paragraph":
            claims.append(_Claim(scope=scope, kind="whole"))
        elif operation.op == "clear-run-properties":
            run = int(args["run_index"])
            claims.extend(
                _Claim(scope=scope, kind="run-property", run=run, prop=str(prop))
                for prop in args["properties"]
            )
        elif operation.op == "set-run-language":
            claims.append(
                _Claim(
                    scope=scope,
                    kind="run-property",
                    run=int(args["run_index"]),
                    prop="lang",
                )
            )
        elif operation.op == "clear-paragraph-properties":
            claims.extend(
                _Claim(scope=scope, kind="paragraph-property", prop=str(prop))
                for prop in args["properties"]
            )
        elif operation.op == "clear-paragraph-numbering":
            claims.append(_Claim(scope=scope, kind="paragraph-property", prop="num_id"))
        elif operation.op == "replace-paragraph-text":
            claims.extend(
                _Claim(scope=scope, kind="text", span=(int(span["start"]), int(span["end"])))
                for span in args["spans"]
            )
    return claims


def _collides(left: _Claim, right: _Claim) -> bool:
    """Whether two claims cannot both be honoured."""
    if left.scope != right.scope:
        return False  # different paragraph, or different style
    if left.kind == "whole" or right.kind == "whole":
        # Removing the thing conflicts with every other edit to it.
        return True
    if left.kind != right.kind:
        return False
    if left.kind == "text":
        # Half-open, so [0, 4) and [4, 9) are adjacent rather than
        # overlapping — two rules editing neighbouring spans of the same
        # paragraph compose perfectly well.
        return left.span[0] < right.span[1] and right.span[0] < left.span[1]
    return (left.run, left.prop) == (right.run, right.prop)


def _describe(claim: _Claim) -> str:
    """Render a claim for a conflict message."""
    kind, name = claim.scope
    where = f"style {name}" if kind == "style" else f"paragraph {name}"
    if claim.kind == "whole":
        return f"{where} is removed"
    if claim.kind == "text":
        return f"{where}, characters {claim.span[0]}-{claim.span[1]}"
    if claim.kind == "run-property":
        return f"{where}, run {claim.run}, {claim.prop}"
    return f"{where}, {claim.prop}"


__all__ = ["FixConflict", "FixPlan", "InvalidFixError", "PlannedFix", "plan_fixes"]
