"""Populate every control in a built form with sample values.

Companion to :mod:`docx_plus.examples.build_form`. Demonstrates the
read/write surface: :func:`read_controls` to discover what's there,
:func:`set_control_value` to fill, then :func:`clear_control` to round-trip
a reset.

Usage::

    python -m docx_plus.examples.populate_form              # builds form.docx then fills
    python -m docx_plus.examples.populate_form form.docx    # fills existing form.docx in place
    python -m docx_plus.examples.populate_form form.docx filled.docx

The no-arg form runs :func:`docx_plus.examples.build_form.build_onboarding_form`
into a tmp dir first, so the example chain works without manual setup.
"""

from __future__ import annotations

import sys
import tempfile
from datetime import datetime
from pathlib import Path

from docx import Document

from docx_plus.controls import (
    ControlValue,
    clear_control,
    read_controls,
    set_control_value,
)
from docx_plus.examples.build_form import build_onboarding_form

SAMPLE_VALUES: dict[str, str | bool | datetime] = {
    "full_name": "Ada Lovelace",
    "display_name": "Ada",
    "start_date": datetime(2026, 6, 1),
    "department": "ENG",  # value, not display — see _set_with_fallback
    "office": "London",
    "remote_first": True,
    "equipment_notes": '27" monitor; standing desk preferred.',
}


def _print_controls(label: str, controls: dict[str, ControlValue]) -> None:
    """Pretty-print a control snapshot."""
    print(f"# {label}")
    if not controls:
        print("  (no controls)")
        return
    name_w = max(len(t) for t in controls)
    for tag, ctrl in controls.items():
        marker = " (placeholder)" if ctrl.is_placeholder else ""
        print(f"  {tag:<{name_w}}  [{ctrl.control_type:<8}]  {ctrl.value!r}{marker}")


def populate_form(in_path: Path, out_path: Path) -> None:
    """Fill every known control in ``in_path``, save to ``out_path``."""
    doc = Document(str(in_path))

    before = read_controls(doc)
    _print_controls("before:", before)
    print()

    for tag, value in SAMPLE_VALUES.items():
        if tag not in before:
            print(f"  skip {tag} (not present in form)")
            continue
        set_control_value(doc, tag, value)

    # Demonstrate clear_control: reset one field and re-read so the output
    # shows a mixed filled/placeholder state.
    clear_control(doc, "display_name")

    doc.save(str(out_path))
    print(f"# wrote: {out_path}")
    print()

    after = read_controls(Document(str(out_path)))
    _print_controls("after:", after)


def main(argv: list[str] | None = None) -> int:
    """Entry point.

    - 0 args: build a form in a tmp dir, then populate into cwd/filled.docx.
    - 1 arg:  populate the given form in-place.
    - 2 args: read first arg, write second.
    """
    args = list(sys.argv[1:] if argv is None else argv)

    if len(args) == 0:
        tmp_dir = Path(tempfile.mkdtemp(prefix="docxplus_populate_"))
        in_path = build_onboarding_form(tmp_dir / "form.docx")
        out_path = Path.cwd() / "filled.docx"
        print(f"# built form: {in_path}")
        print()
    elif len(args) == 1:
        in_path = Path(args[0]).expanduser().resolve()
        out_path = in_path
    elif len(args) == 2:
        in_path = Path(args[0]).expanduser().resolve()
        out_path = Path(args[1]).expanduser().resolve()
    else:
        print(
            "usage: python -m docx_plus.examples.populate_form [input.docx [output.docx]]",
            file=sys.stderr,
        )
        return 2

    if not in_path.exists():
        print(f"error: {in_path} not found", file=sys.stderr)
        return 1

    populate_form(in_path, out_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
