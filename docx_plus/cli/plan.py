"""``docx-plus plan`` — show what repairing a document would change.

Wraps :func:`docx_plus.lint.lint` followed by
:func:`docx_plus.lint.plan_fixes`. A pure read like ``lint``, and for the
same reason it takes ``--json`` and no ``-o/--output``: **this release
applies nothing**. The command's job is to make the repair inspectable
before anything can perform it.

The four sections of the output are the four things that can happen to a
finding — it becomes an edit, it is withheld for changing content, it
loses a collision with another edit, or nobody knows how to fix it. Every
finding lands in exactly one, so the report accounts for the whole audit
rather than quietly listing only the good news.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx_plus.cli._io import add_lint_options, dump_json, load_document, resolve_profile
from docx_plus.lint import lint, plan_fixes

if TYPE_CHECKING:
    import argparse

    from docx_plus.lint import FixPlan, PlannedFix


_SAFETY_MARK = {"safe": "safe ", "review": "check", "destructive": "DROP "}


def register(subparsers: argparse._SubParsersAction[argparse.ArgumentParser]) -> None:
    """Add the ``plan`` subparser."""
    parser = subparsers.add_parser(
        "plan",
        help="show the edits that would repair a document",
        description=(
            "Audit a document and describe the repair: which edits would be made, "
            "in what order, which are withheld for changing content, and which "
            "collide. Read-only - nothing is modified and nothing is applied."
        ),
    )
    parser.add_argument("file", help="path to the .docx file to plan against")
    add_lint_options(parser)
    parser.add_argument(
        "--allow-content",
        action="store_true",
        help=(
            "include edits that change what the document contains (deleting a "
            "paragraph or a style) rather than only how it looks"
        ),
    )
    parser.add_argument(
        "--json",
        dest="as_json",
        action="store_true",
        help="emit structured JSON instead of text",
    )
    parser.set_defaults(func=cmd_plan)


def _print_fix(planned: PlannedFix, number: int, width: int) -> None:
    """Print one planned fix and the operations it would run."""
    mark = _SAFETY_MARK[planned.safety]
    where = planned.finding.location.describe()
    print(f"{number:>{width}}. [{mark}] {where}  {planned.rule}")
    print(f"{'':>{width}}    {planned.fix.summary}")
    for operation in planned.operations:
        print(f"{'':>{width}}    - {operation.op} {_render_args(operation)}")


def _render_args(operation: object) -> str:
    """Render an operation's arguments compactly, ASCII-only for a cp1252 console."""
    args = getattr(operation, "args", {})
    parts = []
    for key, value in args.items():
        if key == "spans":
            rendered = "; ".join(
                f"{span['start']}-{span['end']}->{span['replacement']!r}" for span in value
            )
            parts.append(f"spans=[{rendered}]")
        elif isinstance(value, list):
            parts.append(f"{key}={','.join(str(item) for item in value)}")
        else:
            parts.append(f"{key}={value}")
    return " ".join(parts)


def _print_plan(plan: FixPlan, *, allow_content: bool) -> None:
    """Print the plan as four sections, then a one-line tally."""
    if plan.fixes:
        width = len(str(len(plan.fixes)))
        print(f"{len(plan.fixes)} edit(s), in the order they would be applied:\n")
        for number, planned in enumerate(plan.fixes, start=1):
            _print_fix(planned, number, width)
        print()
    else:
        print("No edits.\n")

    if plan.deferred:
        print(f"{len(plan.deferred)} withheld - these change what the document contains:")
        for planned in plan.deferred:
            print(f"  {planned.finding.location.describe()}  {planned.rule}")
            print(f"    {planned.fix.summary}")
        if not allow_content:
            print("  (re-run with --allow-content to include them)")
        print()

    if plan.conflicts:
        print(f"{len(plan.conflicts)} conflict(s) - the later edit was dropped:")
        for conflict in plan.conflicts:
            print(f"  {conflict.dropped.rule} dropped; {conflict.kept.rule} claimed it first")
            print(f"    both target {conflict.reason}")
        print()

    if plan.unfixable:
        by_rule: dict[str, int] = {}
        for finding in plan.unfixable:
            by_rule[finding.rule] = by_rule.get(finding.rule, 0) + 1
        print(f"{len(plan.unfixable)} finding(s) with no known repair:")
        for rule_id, count in sorted(by_rule.items()):
            print(f"  {count:>4}  {rule_id}")
        print()

    print(
        f"{len(plan.fixes)} to apply, {len(plan.deferred)} withheld, "
        f"{len(plan.conflicts)} dropped, {len(plan.unfixable)} unfixable."
    )


def cmd_plan(args: argparse.Namespace) -> int:
    """Handle ``docx-plus plan``.

    Returns:
        ``0`` when there is nothing to do, ``1`` when the plan holds any
        edit — applied or withheld — so the command gates a pipeline the
        same way ``lint`` does. Findings nobody can repair do not fail the
        gate: they are not something a repair pass would have fixed.
    """
    doc = load_document(args.file)
    findings = lint(
        doc,
        select=args.select,
        exclude=args.exclude,
        include_tables=args.include_tables,
        profile=resolve_profile(args),
    )
    plan = plan_fixes(findings, allow_content=args.allow_content)

    if args.as_json:
        dump_json(plan.to_dict())
    else:
        _print_plan(plan, allow_content=args.allow_content)

    return 1 if plan.fixes or plan.deferred else 0
