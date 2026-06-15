"""The ``docx-plus`` command-line interface.

A thin shell over the library: each subcommand wraps one existing, tested
function. :func:`main` is the console entry point (registered in
``pyproject.toml`` as ``docx-plus = "docx_plus.cli:main"``) and is also runnable
as ``python -m docx_plus.cli``.

Commands:

- ``inspect`` — dump effective formatting per paragraph
  (:func:`docx_plus.styles.resolve_effective_formatting`).
- ``restyle`` — remap styles onto canonical ids
  (:func:`docx_plus.styles.remap_styles`).
- ``controls`` — list / set / clear content-control values
  (:mod:`docx_plus.controls`).

Read commands take ``--json``; mutating commands require ``-o/--output`` (or an
explicit ``--in-place``) so the input is never overwritten by accident.
"""

from __future__ import annotations

import argparse
import sys

from docx_plus import __version__
from docx_plus.cli import controls, inspect, restyle
from docx_plus.cli._io import CliError
from docx_plus.core import DocxPlusError


def build_parser() -> argparse.ArgumentParser:
    """Construct the top-level argument parser with every subcommand registered."""
    parser = argparse.ArgumentParser(
        prog="docx-plus",
        description="OOXML-level command-line tools for .docx files.",
    )
    parser.add_argument("--version", action="version", version=f"docx-plus {__version__}")
    subparsers = parser.add_subparsers(dest="command", metavar="<command>")
    for module in (inspect, restyle, controls):
        module.register(subparsers)
    return parser


def main(argv: list[str] | None = None) -> int:
    """Entry point: parse ``argv`` (default ``sys.argv[1:]``) and dispatch.

    Returns:
        ``0`` on success, ``1`` for a handled library/CLI error, ``2`` when no
        command was given (argparse uses ``2`` for usage errors too).
    """
    parser = build_parser()
    args = parser.parse_args(argv)
    func = getattr(args, "func", None)
    if func is None:
        parser.print_help()
        return 2
    try:
        return int(func(args))
    except DocxPlusError as exc:
        # CliError and every typed library error land here.
        kind = "error" if isinstance(exc, CliError) else f"{type(exc).__name__}"
        print(f"error: {exc}" if isinstance(exc, CliError) else f"{kind}: {exc}", file=sys.stderr)
        return 1


__all__ = ["build_parser", "main"]
