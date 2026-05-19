"""Build a fillable employee onboarding form using :class:`FormBuilder`.

Demonstrates every Phase 4 control type (text, multiline text, dropdown,
combobox, date picker, checkbox) plus the Phase 5 ``protect_document`` call
that turns the result into a "fillable form" — readers can edit the
controls in Word but not the surrounding template text.

Usage::

    python -m docx_plus.examples.build_form               # writes ./form.docx
    python -m docx_plus.examples.build_form path/out.docx

The form's controls all use stable ``tag=`` values, so the companion
:mod:`docx_plus.examples.populate_form` script can fill them in
deterministically.
"""

from __future__ import annotations

import sys
from pathlib import Path

from docx_plus.controls import FormBuilder
from docx_plus.protection import protect_document


def build_onboarding_form(out_path: Path) -> Path:
    """Build the onboarding form and save to ``out_path``. Returns the path."""
    fb = FormBuilder()

    fb.doc.add_heading("New employee onboarding form", level=1)
    fb.doc.add_paragraph(
        "Please fill in every field. Tab from field to field; the document "
        "is protected against accidental edits outside the controls."
    )

    # --- Identity ----------------------------------------------------------
    fb.doc.add_heading("Identity", level=2)

    para = fb.doc.add_paragraph("Full legal name: ")
    fb.add_text_control(
        para,
        tag="full_name",
        alias="Full name",
        placeholder="Type your full name",
    )

    para = fb.doc.add_paragraph("Preferred display name: ")
    fb.add_text_control(
        para,
        tag="display_name",
        placeholder="(optional)",
    )

    para = fb.doc.add_paragraph("Start date: ")
    fb.add_date_picker(
        para,
        tag="start_date",
        alias="Start date",
        date_format="M/d/yyyy",
    )

    # --- Role --------------------------------------------------------------
    fb.doc.add_heading("Role", level=2)

    para = fb.doc.add_paragraph("Department: ")
    fb.add_dropdown(
        para,
        tag="department",
        alias="Department",
        items=[
            ("Engineering", "ENG"),
            ("Design", "DES"),
            ("Operations", "OPS"),
            ("People", "PPL"),
        ],
        placeholder="Choose a department",
    )

    para = fb.doc.add_paragraph("Office (combobox — type if not listed): ")
    fb.add_dropdown(
        para,
        tag="office",
        items=["New York", "London", "Singapore", "Remote"],
        editable=True,  # combobox: accepts freeform values
        placeholder="Pick or type",
    )

    # --- Preferences -------------------------------------------------------
    fb.doc.add_heading("Preferences", level=2)

    para = fb.doc.add_paragraph("Remote-first? ")
    fb.add_checkbox(para, tag="remote_first", checked=False)

    para = fb.doc.add_paragraph("Equipment notes: ")
    fb.add_text_control(
        para,
        tag="equipment_notes",
        alias="Equipment notes",
        placeholder="Any accessibility needs, monitor size, etc.",
        multiline=True,
    )

    # Lock the surrounding text so only the controls accept input.
    protect_document(fb.doc, mode="forms")

    return Path(fb.save(str(out_path)))


def main(argv: list[str] | None = None) -> int:
    """Entry point. One optional positional arg: the output docx path."""
    args = list(sys.argv[1:] if argv is None else argv)
    if len(args) > 1:
        print(
            "usage: python -m docx_plus.examples.build_form [output.docx]",
            file=sys.stderr,
        )
        return 2

    out_path = Path(args[0]).expanduser().resolve() if args else Path.cwd() / "form.docx"

    written = build_onboarding_form(out_path)
    print(f"# wrote: {written}")
    print("# 7 controls: full_name, display_name, start_date, department,")
    print("#             office, remote_first, equipment_notes")
    print("# protection: forms (only controls accept input)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
