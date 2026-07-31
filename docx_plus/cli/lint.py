"""``docx-plus lint`` — audit a document for formatting defects.

Wraps :func:`docx_plus.lint.lint`. A pure read, so it takes ``--json`` and
never writes: applying fixes is a separate command, deliberately not in this
release.

``--rule`` / ``--exclude`` take a rule id or a tag; naming a tag also enables
that cluster's off-by-default rules. ``--list-rules`` prints the catalogue
without needing a document.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from docx_plus.cli._io import add_lint_options, dump_json, load_document, resolve_profile
from docx_plus.lint import all_rules, lint

if TYPE_CHECKING:
    import argparse

    from docx_plus.lint import Finding


_SEVERITY_MARK = {"error": "E", "warning": "W", "info": "i"}


def register(subparsers: argparse._SubParsersAction[argparse.ArgumentParser]) -> None:
    """Add the ``lint`` subparser."""
    parser = subparsers.add_parser(
        "lint",
        help="audit a document for formatting defects",
        description=(
            "Report formatting defects: direct formatting fighting the styles, "
            "outline problems, hand-typed lists, whitespace used as layout. "
            "Read-only — nothing is modified."
        ),
    )
    parser.add_argument("file", nargs="?", help="path to the .docx file to audit")
    add_lint_options(parser)
    parser.add_argument(
        "--list-rules",
        action="store_true",
        help="print the rule catalogue and exit (no document needed)",
    )
    parser.add_argument(
        "--json",
        dest="as_json",
        action="store_true",
        help="emit structured JSON instead of text",
    )
    parser.set_defaults(func=cmd_lint)


def _finding_json(finding: Finding) -> dict[str, Any]:
    """Build the JSON record for one finding."""
    return {
        "rule": finding.rule,
        "kind": finding.kind,
        "severity": finding.severity,
        "message": finding.message,
        "location": {
            "paragraph_index": finding.location.paragraph_index,
            "run_index": finding.location.run_index,
            "style_id": finding.location.style_id,
            "excerpt": finding.location.excerpt,
        },
        "observed": finding.observed,
        "expected": finding.expected,
        "fixable": finding.fixable,
        "adds_content": finding.adds_content,
    }


def _print_findings(findings: list[Finding]) -> None:
    """Print findings as an aligned table, then a per-severity summary."""
    if not findings:
        print("No findings.")
        return

    where_w = max(len(f.location.describe()) for f in findings)
    rule_w = max(len(f.rule) for f in findings)

    for finding in findings:
        mark = _SEVERITY_MARK[finding.severity]
        where = finding.location.describe()
        print(f"{mark} {where:<{where_w}}  {finding.rule:<{rule_w}}  {finding.message}")
        if finding.location.excerpt:
            # Quoted, so leading and trailing whitespace is visible — several
            # rules are precisely about whitespace nobody can see.
            print(f'{"":<{where_w + 2}}  {"":<{rule_w}}  > "{finding.location.excerpt}"')

    counts = {level: sum(1 for f in findings if f.severity == level) for level in _SEVERITY_MARK}
    summary = ", ".join(f"{n} {level}" for level, n in counts.items() if n)
    print(f"\n{len(findings)} finding{'s' if len(findings) != 1 else ''} ({summary}).")


def _print_rules() -> None:
    """Print the rule catalogue as an aligned table."""
    rules = all_rules()
    id_w = max(len(r.id) for r in rules)
    kind_w = max(len(r.kind) for r in rules)

    for registered in rules:
        state = "on " if registered.default_on else "off"
        tags = ",".join(sorted(registered.tags))
        print(
            f"{registered.id:<{id_w}}  {registered.kind:<{kind_w}}  "
            f"{registered.severity:<7}  {state}  [{tags}]"
        )
        print(f"{'':<{id_w}}  {registered.description}")

    on = sum(1 for r in rules if r.default_on)
    print(f"\n{len(rules)} rules, {on} on by default.")


def _rules_json() -> list[dict[str, Any]]:
    """Build the JSON record for the rule catalogue."""
    return [
        {
            "id": r.id,
            "kind": r.kind,
            "severity": r.severity,
            "description": r.description,
            "tags": sorted(r.tags),
            "default_on": r.default_on,
        }
        for r in all_rules()
    ]


def cmd_lint(args: argparse.Namespace) -> int:
    """Handle ``docx-plus lint``.

    Returns:
        ``0`` when the document is clean or the catalogue was printed, ``1``
        when any finding was reported — so the command is usable as a CI
        gate.
    """
    if args.list_rules:
        if args.as_json:
            dump_json(_rules_json())
        else:
            _print_rules()
        return 0

    if not args.file:
        # argparse cannot express "required unless --list-rules".
        from docx_plus.cli._io import CliError

        raise CliError("a FILE is required unless --list-rules is given")

    doc = load_document(args.file)
    findings = lint(
        doc,
        select=args.select,
        exclude=args.exclude,
        include_tables=args.include_tables,
        profile=resolve_profile(args),
    )

    if args.as_json:
        dump_json([_finding_json(f) for f in findings])
    else:
        _print_findings(findings)

    return 1 if findings else 0
