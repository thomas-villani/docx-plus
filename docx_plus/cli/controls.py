"""``docx-plus controls`` — list, set, or clear content-control values.

Wraps :func:`docx_plus.controls.list_controls`,
:func:`~docx_plus.controls.set_control_value`, and
:func:`~docx_plus.controls.clear_control`. ``list`` is read-only; ``set`` and
``clear`` mutate and therefore require ``-o/--output`` (or ``--in-place``).

``list`` reports every control, including the untagged and empty-tag ones Word
writes by default, so ``--tag`` is not always enough to address one. Every
listing therefore prints the control's ``w:id``, which ``--control-id`` accepts
as an unambiguous alternative to ``--tag``.

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
    ControlValue,
    clear_control,
    list_controls,
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
        help="label controls by tag (default) or alias; controls with neither "
        "are labelled #INDEX and are still listed",
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
    _add_selector_args(set_p)
    set_p.add_argument(
        "--value", required=True, help="the new value (coerced to the control's type)"
    )
    _add_output_args(set_p)
    set_p.set_defaults(func=cmd_set)

    clear_p = actions.add_parser("clear", help="reset a control to its placeholder state")
    clear_p.add_argument("file", help="path to the source .docx file")
    _add_selector_args(clear_p)
    _add_output_args(clear_p)
    clear_p.set_defaults(func=cmd_clear)


def _add_selector_args(parser: argparse.ArgumentParser) -> None:
    """Add the shared --tag / --control-id target selectors to a mutating action.

    Word writes most controls with an empty ``w:tag``, so a tag cannot be
    required and cannot be assumed unique. Exactly one selector must be given;
    ``--control-id`` is the one that always resolves.
    """
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--tag", help="the control's w:tag value (must match exactly one)")
    group.add_argument(
        "--control-id",
        type=int,
        help="the control's w:id value, as shown by 'controls list'",
    )


def _add_output_args(parser: argparse.ArgumentParser) -> None:
    """Add the shared -o/--output and --in-place options to a mutating action."""
    parser.add_argument("-o", "--output", default=None, help="path to write the result")
    parser.add_argument(
        "--in-place",
        action="store_true",
        help="overwrite the input file instead of requiring -o/--output",
    )


def _label(cv: ControlValue, by: str) -> str:
    """Display label for a control: its tag/alias, or ``#INDEX`` when it has none."""
    key = cv.tag if by == "tag" else cv.alias
    return key if key else f"#{cv.index}"


def cmd_list(args: argparse.Namespace) -> int:
    """Handle ``docx-plus controls list``."""
    doc = load_document(args.file)
    controls = list_controls(doc)
    if args.as_json:
        dump_json(
            [
                {
                    "key": _label(cv, args.by),
                    "index": cv.index,
                    "tag": cv.tag,
                    "alias": cv.alias,
                    "control_id": cv.control_id,
                    "control_type": cv.control_type,
                    "location": cv.location,
                    "value": cv.value,
                    "is_placeholder": cv.is_placeholder,
                }
                for cv in controls
            ]
        )
        return 0

    if not controls:
        print("(no content controls)")
        return 0
    for cv in controls:
        value = "(placeholder)" if cv.is_placeholder else repr(cv.value)
        alias = f" alias={cv.alias!r}" if cv.alias and args.by != "alias" else ""
        ident = f" id={cv.control_id}" if cv.control_id is not None else ""
        where = "" if cv.location == "body" else f" in {cv.location}"
        print(f"{_label(cv, args.by)}: {cv.control_type}{alias}{ident}{where} = {value}")
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


def _select(doc: DocumentObj, args: argparse.Namespace) -> ControlValue:
    """Resolve ``--tag`` / ``--control-id`` to exactly one control.

    Reports ambiguity rather than picking a match, because ``--tag ""`` matches
    every untagged control in a typical Word form and writing to an arbitrary
    one of them looks like success.
    """
    if args.control_id is not None:
        matches = [c for c in list_controls(doc) if c.control_id == args.control_id]
        if not matches:
            raise CliError(f"no control with id {args.control_id}")
        return matches[0]

    matches = [c for c in list_controls(doc) if c.tag == args.tag]
    if not matches:
        raise CliError(f"no control with tag {args.tag!r}")
    if len(matches) > 1:
        ids = ", ".join(str(c.control_id) for c in matches)
        raise CliError(
            f"tag {args.tag!r} matches {len(matches)} controls (ids {ids}); "
            f"pass --control-id to choose one"
        )
    return matches[0]


def _describe(target: ControlValue) -> str:
    """How to refer to a control in command output."""
    return repr(target.tag) if target.tag else f"id {target.control_id}"


def _selector(target: ControlValue) -> tuple[str | None, int | None]:
    """The ``(tag, control_id)`` pair to hand the write API for ``target``.

    Prefers ``w:id`` because it survives a repeated tag, but falls back to the
    tag for the rare control that carries no usable id.
    """
    if target.control_id is not None:
        return None, target.control_id
    return target.tag, None


def cmd_set(args: argparse.Namespace) -> int:
    """Handle ``docx-plus controls set``."""
    out_path = resolve_output(args)
    doc = load_document(args.file)
    target = _select(doc, args)
    coerced = _coerce_value(args.value, target.control_type)
    tag, control_id = _selector(target)
    set_control_value(doc, tag, coerced, control_id=control_id)
    save_document(doc, out_path)
    print(f"set {_describe(target)} = {coerced!r}; wrote {out_path}")
    return 0


def cmd_clear(args: argparse.Namespace) -> int:
    """Handle ``docx-plus controls clear``."""
    out_path = resolve_output(args)
    doc = load_document(args.file)
    target = _select(doc, args)
    tag, control_id = _selector(target)
    clear_control(doc, tag, control_id=control_id)
    save_document(doc, out_path)
    print(f"cleared {_describe(target)}; wrote {out_path}")
    return 0
