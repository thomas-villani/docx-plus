"""``docx-plus controls`` — list, set, or clear content-control values.

Wraps :func:`docx_plus.controls.read_controls`,
:func:`~docx_plus.controls.set_control_value`, and
:func:`~docx_plus.controls.clear_control`. ``list`` is read-only; ``set`` and
``clear`` mutate and therefore require ``-o/--output`` (or ``--in-place``).

Because the command line only carries strings, ``set`` reads the target
control's type first and coerces the supplied ``--value`` to the Python type the
underlying API requires (``bool`` for checkboxes, :class:`~datetime.datetime`
for dates, ``str`` otherwise).
"""

from __future__ import annotations

from datetime import datetime
from typing import TYPE_CHECKING

from docx_plus.cli._io import (
    CliError,
    dump_json,
    load_document,
    resolve_output,
    save_document,
)
from docx_plus.controls import (
    ControlType,
    clear_control,
    read_controls,
    set_control_value,
)

if TYPE_CHECKING:
    import argparse

    from docx.document import Document as DocumentObj

_TRUE = frozenset({"true", "1", "yes", "on"})
_FALSE = frozenset({"false", "0", "no", "off"})


def register(subparsers: argparse._SubParsersAction[argparse.ArgumentParser]) -> None:
    """Add the ``controls`` subparser and its list/set/clear sub-actions."""
    parser = subparsers.add_parser(
        "controls",
        help="list, set, or clear content-control values",
        description="Inspect and edit content controls (fillable form fields).",
    )
    actions = parser.add_subparsers(dest="action", metavar="{list,set,clear}")
    actions.required = True

    list_p = actions.add_parser("list", help="list every control and its value")
    list_p.add_argument("file", help="path to the .docx file")
    list_p.add_argument(
        "--by",
        choices=("tag", "alias"),
        default="tag",
        help="key controls by tag (default) or alias",
    )
    list_p.add_argument(
        "--json",
        dest="as_json",
        action="store_true",
        help="emit structured JSON instead of text",
    )
    list_p.set_defaults(func=cmd_list)

    set_p = actions.add_parser("set", help="set a control's value")
    set_p.add_argument("file", help="path to the source .docx file")
    set_p.add_argument("--tag", required=True, help="the control's w:tag value")
    set_p.add_argument(
        "--value", required=True, help="the new value (coerced to the control's type)"
    )
    _add_output_args(set_p)
    set_p.set_defaults(func=cmd_set)

    clear_p = actions.add_parser("clear", help="reset a control to its placeholder state")
    clear_p.add_argument("file", help="path to the source .docx file")
    clear_p.add_argument("--tag", required=True, help="the control's w:tag value")
    _add_output_args(clear_p)
    clear_p.set_defaults(func=cmd_clear)


def _add_output_args(parser: argparse.ArgumentParser) -> None:
    """Add the shared -o/--output and --in-place options to a mutating action."""
    parser.add_argument("-o", "--output", default=None, help="path to write the result")
    parser.add_argument(
        "--in-place",
        action="store_true",
        help="overwrite the input file instead of requiring -o/--output",
    )


def cmd_list(args: argparse.Namespace) -> int:
    """Handle ``docx-plus controls list``."""
    doc = load_document(args.file)
    controls = read_controls(doc, by=args.by)
    if args.as_json:
        dump_json(
            [
                {
                    "key": key,
                    "tag": cv.tag,
                    "alias": cv.alias,
                    "control_type": cv.control_type,
                    "value": cv.value,
                    "is_placeholder": cv.is_placeholder,
                }
                for key, cv in controls.items()
            ]
        )
        return 0

    if not controls:
        print("(no content controls)")
        return 0
    for key, cv in controls.items():
        value = "(placeholder)" if cv.is_placeholder else repr(cv.value)
        alias = f" alias={cv.alias!r}" if cv.alias else ""
        print(f"{key}: {cv.control_type}{alias} = {value}")
    return 0


def _coerce_value(raw: str, control_type: ControlType) -> str | bool | datetime:
    """Coerce a command-line string to the Python type the control requires."""
    if control_type == "checkbox":
        lowered = raw.strip().lower()
        if lowered in _TRUE:
            return True
        if lowered in _FALSE:
            return False
        raise CliError(f"checkbox value must be true/false, got {raw!r}")
    if control_type == "date":
        try:
            return datetime.fromisoformat(raw)
        except ValueError as exc:
            raise CliError(f"date value must be ISO 8601, got {raw!r}") from exc
    return raw


def _control_type(doc: DocumentObj, tag: str) -> ControlType:
    """Look up the type of the control identified by ``tag``."""
    controls = read_controls(doc)
    if tag not in controls:
        raise CliError(f"no control with tag {tag!r}")
    return controls[tag].control_type


def cmd_set(args: argparse.Namespace) -> int:
    """Handle ``docx-plus controls set``."""
    out_path = resolve_output(args)
    doc = load_document(args.file)
    coerced = _coerce_value(args.value, _control_type(doc, args.tag))
    set_control_value(doc, args.tag, coerced)
    save_document(doc, out_path)
    print(f"set {args.tag!r} = {coerced!r}; wrote {out_path}")
    return 0


def cmd_clear(args: argparse.Namespace) -> int:
    """Handle ``docx-plus controls clear``."""
    out_path = resolve_output(args)
    doc = load_document(args.file)
    clear_control(doc, args.tag)
    save_document(doc, out_path)
    print(f"cleared {args.tag!r}; wrote {out_path}")
    return 0
