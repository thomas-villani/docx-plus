"""``docx-plus inspect`` — dump effective formatting per paragraph.

Wraps :func:`docx_plus.styles.resolve_effective_formatting`. For each paragraph
in the document it prints the fields the cascade actually resolved; with
``--provenance`` each field is annotated with the cascade layer that produced it.
``--json`` emits the same data as a structured list for piping into other tools.
"""

from __future__ import annotations

from dataclasses import fields as dataclass_fields
from typing import TYPE_CHECKING, Any

from docx_plus.cli._io import dump_json, load_document
from docx_plus.styles import (
    FormattingSource,
    ResolvedFormatting,
    resolve_effective_formatting,
)

if TYPE_CHECKING:
    import argparse

# style_id/style_name head their own line; partial/provenance are meta-fields.
_META_FIELDS = frozenset({"style_id", "style_name", "partial", "provenance"})


def register(subparsers: argparse._SubParsersAction[argparse.ArgumentParser]) -> None:
    """Add the ``inspect`` subparser."""
    parser = subparsers.add_parser(
        "inspect",
        help="dump the effective formatting of each paragraph",
        description="Resolve and print the effective formatting for every paragraph in a document.",
    )
    parser.add_argument("file", help="path to the .docx file to inspect")
    parser.add_argument(
        "--provenance",
        action="store_true",
        help="annotate each field with the cascade layer that set it",
    )
    parser.add_argument(
        "--json",
        dest="as_json",
        action="store_true",
        help="emit structured JSON instead of text",
    )
    parser.set_defaults(func=cmd_inspect)


def _format_source(src: FormattingSource) -> str:
    """Render a FormattingSource as a one-line suffix (mirrors the example)."""
    parts: list[str] = [src.layer]
    if src.style_id is not None:
        parts.append(f": {src.style_id}")
    extras: list[str] = []
    if src.chain_depth is not None and src.chain_depth > 0:
        extras.append(f"chain_depth={src.chain_depth}")
    if src.is_toggle_resolved:
        extras.append("toggle XOR")
    if extras:
        parts.append(f" ({', '.join(extras)})")
    return "".join(parts)


def _set_fields(resolved: ResolvedFormatting) -> list[tuple[str, Any]]:
    """Return the (name, value) pairs the cascade actually set, in field order."""
    out: list[tuple[str, Any]] = []
    for f in dataclass_fields(resolved):
        if f.name in _META_FIELDS:
            continue
        value = getattr(resolved, f.name)
        if value is not None:
            out.append((f.name, value))
    return out


def _print_paragraph(index: int, text: str, resolved: ResolvedFormatting) -> None:
    """Print one paragraph's resolved formatting as text."""
    snippet = text if len(text) <= 60 else text[:57] + "..."
    print(f'[{index}] "{snippet}"')

    style_label = resolved.style_id or "(no style)"
    if resolved.style_name and resolved.style_name != resolved.style_id:
        style_label = f"{resolved.style_id} ({resolved.style_name})"
    print(f"    style: {style_label}")

    provenance = resolved.provenance or {}
    rows = _set_fields(resolved)
    if not rows:
        print("    (no fields set)")
    else:
        name_w = max(len(name) for name, _ in rows)
        value_w = max(len(repr(v)) for _, v in rows)
        for name, value in rows:
            src = provenance.get(name)
            suffix = f"  <- {_format_source(src)}" if src is not None else ""
            print(f"    {name:<{name_w}} : {repr(value):<{value_w}}{suffix}")

    if resolved.partial:
        print("    (partial: theme part missing or malformed)")
    print()


def _paragraph_json(index: int, text: str, resolved: ResolvedFormatting) -> dict[str, Any]:
    """Build the JSON record for one paragraph."""
    provenance = resolved.provenance or {}
    fields = {name: value for name, value in _set_fields(resolved)}
    record: dict[str, Any] = {
        "index": index,
        "text": text,
        "style_id": resolved.style_id,
        "style_name": resolved.style_name,
        "partial": resolved.partial,
        "fields": fields,
    }
    if resolved.provenance is not None:
        record["provenance"] = {
            name: _format_source(provenance[name]) for name in fields if name in provenance
        }
    return record


def cmd_inspect(args: argparse.Namespace) -> int:
    """Handle ``docx-plus inspect``."""
    doc = load_document(args.file)
    records: list[dict[str, Any]] = []
    for index, paragraph in enumerate(doc.paragraphs, start=1):
        text = paragraph.text
        resolved = resolve_effective_formatting(paragraph, include_provenance=args.provenance)
        if args.as_json:
            records.append(_paragraph_json(index, text, resolved))
        else:
            _print_paragraph(index, text or "(empty paragraph)", resolved)
    if args.as_json:
        dump_json(records)
    return 0
