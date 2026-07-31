"""Shared plumbing for the ``docx-plus`` command-line interface.

Every command loads a document, optionally mutates it, and writes results to
stdout (or a new ``.docx``). The helpers here centralize the three concerns the
commands share: opening a document with a friendly error, resolving where a
mutating command may write, and emitting JSON. :class:`CliError` is the one
exception type the dispatcher in :mod:`docx_plus.cli` catches and renders.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import TYPE_CHECKING, Any

from docx import Document

from docx_plus.core import DocxPlusError
from docx_plus.lint import DEFAULT_PROFILE_NAME, Profile

if TYPE_CHECKING:
    import argparse

    from docx.document import Document as DocumentObj


class CliError(DocxPlusError, ValueError):
    """A user-facing CLI failure (bad path, missing output, un-coercible value).

    Raised by command handlers for conditions the user can fix. The dispatcher
    catches it (and any other :class:`~docx_plus.core.DocxPlusError`), prints
    ``error: <message>`` to stderr, and exits non-zero. Subclasses
    ``ValueError`` so ordinary ``except ValueError`` clauses still catch it.
    """


def load_document(path_str: str) -> DocumentObj:
    """Open ``path_str`` as a python-docx document.

    Args:
        path_str: Path to a ``.docx`` file (``~`` is expanded).

    Returns:
        The opened :class:`docx.document.Document`.

    Raises:
        CliError: If the path does not exist or is not a directory entry that
            python-docx can open.
    """
    path = Path(path_str).expanduser()
    if not path.is_file():
        raise CliError(f"{path} not found")
    try:
        return Document(str(path))
    except Exception as exc:  # noqa: BLE001 - re-raise as a clean CLI error
        raise CliError(f"could not open {path}: {exc}") from exc


def resolve_output(args: argparse.Namespace) -> Path:
    """Resolve where a mutating command should write its result.

    Args:
        args: The parsed namespace. Reads ``args.output`` (``-o/--output``),
            ``args.in_place`` (``--in-place``), and ``args.file`` (input path).

    Returns:
        The output path: ``--output`` if given, else the input path when
        ``--in-place`` is set.

    Raises:
        CliError: If neither ``--output`` nor ``--in-place`` was supplied.
    """
    if args.output is not None:
        return Path(args.output).expanduser()
    if getattr(args, "in_place", False):
        return Path(args.file).expanduser()
    raise CliError("specify -o/--output or --in-place")


def save_document(doc: DocumentObj, path: Path) -> None:
    """Save ``doc`` to ``path``, creating parent directories as needed."""
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))


def dump_json(obj: Any) -> None:
    """Print ``obj`` as indented JSON, stringifying non-JSON types (datetimes)."""
    print(json.dumps(obj, indent=2, default=str))


def add_lint_options(parser: argparse.ArgumentParser) -> None:
    """Add the rule-selection and profile options ``lint`` and ``plan`` share.

    Both commands answer the same question about the same document — one
    reports it, the other says what a repair would do — so they take the
    same selectors. Splitting the definitions would let the two drift.
    """
    parser.add_argument(
        "--rule",
        dest="select",
        action="append",
        metavar="ID|TAG",
        help=(
            "only run this rule id or tag (repeatable). Naming a tag also enables "
            "that cluster's off-by-default rules."
        ),
    )
    parser.add_argument(
        "--exclude",
        dest="exclude",
        action="append",
        metavar="ID|TAG",
        help="skip this rule id or tag (repeatable); applied last",
    )
    parser.add_argument(
        "--no-tables",
        dest="include_tables",
        action="store_false",
        help="skip paragraphs inside table cells",
    )
    # Mutually exclusive: naming a profile and then ignoring every profile
    # is a contradiction, and argparse says so better than a runtime check.
    profile_group = parser.add_mutually_exclusive_group()
    profile_group.add_argument(
        "--profile",
        metavar="PATH",
        help=(
            f"lint profile to apply. Without this, {DEFAULT_PROFILE_NAME} is "
            f"looked for beside the document and upwards."
        ),
    )
    profile_group.add_argument(
        "--no-profile",
        action="store_true",
        help="ignore any profile, including a discovered one",
    )


def resolve_profile(args: argparse.Namespace) -> Profile:
    """Load the profile a lint-family command should run under.

    Explicit ``--profile`` wins; otherwise the checked-in profile beside the
    document (or above it) applies, the way every other linter behaves. A
    named path that does not exist is an error rather than a silent fall
    back to discovery — asking for a specific profile and quietly getting a
    different one is the worst of both.
    """
    if args.no_profile:
        # argparse's mutually-exclusive group rejects the combination, so
        # reaching here with both set would be a wiring bug rather than
        # user input. Guard anyway: silently ignoring an explicit --profile
        # is exactly the "asked for one thing, got another" the docstring
        # above rules out, and it hid a nonexistent path entirely.
        if args.profile:
            raise CliError("--profile and --no-profile cannot be combined")
        return Profile()
    if args.profile:
        path = Path(args.profile).expanduser()
        if not path.is_file():
            raise CliError(f"{path} not found")
        return Profile.load(path)
    return Profile.discover(Path(args.file).expanduser())
