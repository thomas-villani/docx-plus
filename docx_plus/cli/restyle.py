"""``docx-plus restyle`` — reconcile a document's styles against canonical ids.

Wraps :func:`docx_plus.styles.remap_styles`. Given one or more target style ids
(``--target``), optional explicit hints (``--map TARGET=EXISTING``), and an
optional ``--create-missing`` flag, it remaps the document's paragraphs and runs
onto the resolved styles, writes the result, and reports the
``target -> resolved-id`` mapping.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx_plus.cli._io import CliError, dump_json, load_document, resolve_output, save_document
from docx_plus.styles import remap_styles

if TYPE_CHECKING:
    import argparse


def register(subparsers: argparse._SubParsersAction[argparse.ArgumentParser]) -> None:
    """Add the ``restyle`` subparser."""
    parser = subparsers.add_parser(
        "restyle",
        help="remap a document's styles onto canonical style ids",
        description="Reconcile a document's styles against a set of canonical "
        "ids, remapping paragraphs and runs onto the resolved styles.",
    )
    parser.add_argument("file", help="path to the source .docx file")
    parser.add_argument(
        "--target",
        action="append",
        metavar="STYLE_ID",
        required=True,
        help="a canonical style id to reconcile (repeatable)",
    )
    parser.add_argument(
        "--map",
        action="append",
        default=[],
        metavar="TARGET=EXISTING",
        dest="mappings",
        help="hint resolving TARGET to an EXISTING style id (repeatable)",
    )
    parser.add_argument(
        "--create-missing",
        action="store_true",
        help="materialize known built-in targets that aren't defined yet",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="path to write the restyled document",
    )
    parser.add_argument(
        "--in-place",
        action="store_true",
        help="overwrite the input file instead of requiring -o/--output",
    )
    parser.add_argument(
        "--json",
        dest="as_json",
        action="store_true",
        help="emit the resolved mapping as JSON",
    )
    parser.set_defaults(func=cmd_restyle)


def _parse_mappings(raw: list[str]) -> dict[str, str]:
    """Parse ``TARGET=EXISTING`` strings into a mapping dict."""
    mapping: dict[str, str] = {}
    for item in raw:
        key, sep, value = item.partition("=")
        if not sep or not key or not value:
            raise CliError(f"invalid --map {item!r}; expected TARGET=EXISTING")
        mapping[key] = value
    return mapping


def cmd_restyle(args: argparse.Namespace) -> int:
    """Handle ``docx-plus restyle``."""
    out_path = resolve_output(args)
    mapping = _parse_mappings(args.mappings)
    doc = load_document(args.file)
    resolved = remap_styles(
        doc,
        targets=args.target,
        mapping=mapping or None,
        create_missing=args.create_missing,
    )
    save_document(doc, out_path)

    if args.as_json:
        dump_json(resolved)
    else:
        print(f"wrote {out_path}")
        if resolved:
            width = max(len(target) for target in resolved)
            for target, style_id in resolved.items():
                print(f"  {target:<{width}} -> {style_id}")
        else:
            print("  (no targets resolved)")
    return 0
