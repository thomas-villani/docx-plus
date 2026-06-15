"""Tests for the ``docx-plus`` command-line interface.

Every test drives the public entry point :func:`docx_plus.cli.main` with an
explicit ``argv`` list and inspects the captured stdout/stderr, the returned
exit code, and any output file. Fixtures are built into ``tmp_path`` so each
test controls the exact styles and control tags it exercises.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from docx import Document

from docx_plus.cli import main
from docx_plus.controls import FormBuilder
from docx_plus.styles import apply_style, ensure_style


@pytest.fixture
def styled_doc(tmp_path: Path) -> Path:
    """A document with a Title paragraph and a plain body paragraph."""
    doc = Document()
    ensure_style(doc, "Title")
    apply_style(doc.add_paragraph("The Title"), "Title")
    doc.add_paragraph("Body text.")
    path = tmp_path / "styled.docx"
    doc.save(str(path))
    return path


@pytest.fixture
def form_doc(tmp_path: Path) -> Path:
    """A form with text / dropdown / date / checkbox controls."""
    fb = FormBuilder()
    p = fb.doc.add_paragraph("Name: ")
    fb.add_text_control(p, tag="name", alias="Full name", placeholder="Type name")
    p = fb.doc.add_paragraph("Dept: ")
    fb.add_dropdown(p, tag="dept", items=["Eng", "Design"])
    p = fb.doc.add_paragraph("Start: ")
    fb.add_date_picker(p, tag="start")
    p = fb.doc.add_paragraph("Subscribed: ")
    fb.add_checkbox(p, tag="subscribed")
    path = tmp_path / "form.docx"
    fb.save(str(path))
    return path


# --------------------------------------------------------------------------
# inspect
# --------------------------------------------------------------------------


def test_inspect_text(styled_doc: Path, capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["inspect", str(styled_doc)]) == 0
    out = capsys.readouterr().out
    assert '[1] "The Title"' in out
    assert "style: Title" in out
    assert "font_size" in out


def test_inspect_provenance_annotates_layer(
    styled_doc: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    assert main(["inspect", str(styled_doc), "--provenance"]) == 0
    out = capsys.readouterr().out
    assert "<- paragraphStyle: Title" in out


def test_inspect_json(styled_doc: Path, capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["inspect", str(styled_doc), "--json", "--provenance"]) == 0
    records = json.loads(capsys.readouterr().out)
    assert records[0]["index"] == 1
    assert records[0]["text"] == "The Title"
    assert records[0]["style_id"] == "Title"
    assert records[0]["fields"]["font_size"] == 26.0
    assert records[0]["provenance"]["font_size"] == "paragraphStyle: Title"


def test_inspect_json_without_provenance_omits_key(
    styled_doc: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    assert main(["inspect", str(styled_doc), "--json"]) == 0
    records = json.loads(capsys.readouterr().out)
    assert "provenance" not in records[0]


def test_inspect_missing_file(capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["inspect", "nope.docx"]) == 1
    assert "not found" in capsys.readouterr().err


# --------------------------------------------------------------------------
# restyle
# --------------------------------------------------------------------------


def test_restyle_writes_output_and_reports_mapping(
    styled_doc: Path, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    out = tmp_path / "restyled.docx"
    assert main(["restyle", str(styled_doc), "--target", "Title", "-o", str(out)]) == 0
    assert out.is_file()
    output = capsys.readouterr().out
    assert f"wrote {out}" in output
    assert "-> Title" in output


def test_restyle_json(styled_doc: Path, tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    out = tmp_path / "restyled.docx"
    code = main(["restyle", str(styled_doc), "--target", "Title", "-o", str(out), "--json"])
    assert code == 0
    mapping = json.loads(capsys.readouterr().out)
    assert mapping == {"Title": "Title"}


def test_restyle_requires_output(styled_doc: Path, capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["restyle", str(styled_doc), "--target", "Title"]) == 1
    assert "specify -o/--output or --in-place" in capsys.readouterr().err


def test_restyle_in_place(styled_doc: Path) -> None:
    assert main(["restyle", str(styled_doc), "--target", "Title", "--in-place"]) == 0
    assert styled_doc.is_file()


def test_restyle_create_missing(styled_doc: Path, tmp_path: Path) -> None:
    out = tmp_path / "restyled.docx"
    code = main(
        [
            "restyle",
            str(styled_doc),
            "--target",
            "Heading1",
            "--create-missing",
            "-o",
            str(out),
        ]
    )
    assert code == 0


def test_restyle_bad_map(
    styled_doc: Path, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    out = tmp_path / "restyled.docx"
    code = main(["restyle", str(styled_doc), "--target", "Title", "--map", "bad", "-o", str(out)])
    assert code == 1
    assert "invalid --map" in capsys.readouterr().err


def test_restyle_map_hint(styled_doc: Path, tmp_path: Path) -> None:
    out = tmp_path / "restyled.docx"
    code = main(
        ["restyle", str(styled_doc), "--target", "Title", "--map", "Title=Title", "-o", str(out)]
    )
    assert code == 0


# --------------------------------------------------------------------------
# controls list
# --------------------------------------------------------------------------


def test_controls_list_text(form_doc: Path, capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["controls", "list", str(form_doc)]) == 0
    out = capsys.readouterr().out
    assert "name: text" in out
    assert "subscribed: checkbox = False" in out


def test_controls_list_json(form_doc: Path, capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["controls", "list", str(form_doc), "--json"]) == 0
    records = json.loads(capsys.readouterr().out)
    by_tag = {r["tag"]: r for r in records}
    assert by_tag["name"]["control_type"] == "text"
    assert by_tag["name"]["is_placeholder"] is True
    assert by_tag["subscribed"]["value"] is False


def test_controls_list_by_alias(form_doc: Path, capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["controls", "list", str(form_doc), "--by", "alias"]) == 0
    out = capsys.readouterr().out
    # Only the text control has an alias; the others are skipped.
    assert "Full name:" in out
    assert "subscribed:" not in out


# --------------------------------------------------------------------------
# controls set / clear
# --------------------------------------------------------------------------


def test_controls_set_text(form_doc: Path, tmp_path: Path) -> None:
    out = tmp_path / "set.docx"
    assert (
        main(["controls", "set", str(form_doc), "--tag", "name", "--value", "Ada", "-o", str(out)])
        == 0
    )
    reread = Document(str(out))
    from docx_plus.controls import read_controls

    assert read_controls(reread)["name"].value == "Ada"


def test_controls_set_checkbox(form_doc: Path, tmp_path: Path) -> None:
    out = tmp_path / "set.docx"
    code = main(
        ["controls", "set", str(form_doc), "--tag", "subscribed", "--value", "yes", "-o", str(out)]
    )
    assert code == 0
    from docx_plus.controls import read_controls

    assert read_controls(Document(str(out)))["subscribed"].value is True


def test_controls_set_date(form_doc: Path, tmp_path: Path) -> None:
    out = tmp_path / "set.docx"
    code = main(
        [
            "controls",
            "set",
            str(form_doc),
            "--tag",
            "start",
            "--value",
            "2026-06-15",
            "-o",
            str(out),
        ]
    )
    assert code == 0
    from docx_plus.controls import read_controls

    value = read_controls(Document(str(out)))["start"].value
    assert value is not None
    assert value.year == 2026 and value.month == 6 and value.day == 15


def test_controls_set_bad_checkbox(
    form_doc: Path, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    out = tmp_path / "set.docx"
    code = main(
        [
            "controls",
            "set",
            str(form_doc),
            "--tag",
            "subscribed",
            "--value",
            "maybe",
            "-o",
            str(out),
        ]
    )
    assert code == 1
    assert "checkbox value must be true/false" in capsys.readouterr().err


def test_controls_set_bad_date(
    form_doc: Path, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    out = tmp_path / "set.docx"
    code = main(
        [
            "controls",
            "set",
            str(form_doc),
            "--tag",
            "start",
            "--value",
            "not-a-date",
            "-o",
            str(out),
        ]
    )
    assert code == 1
    assert "date value must be ISO 8601" in capsys.readouterr().err


def test_controls_set_unknown_tag(
    form_doc: Path, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    out = tmp_path / "set.docx"
    code = main(
        ["controls", "set", str(form_doc), "--tag", "ghost", "--value", "x", "-o", str(out)]
    )
    assert code == 1
    assert "no control with tag 'ghost'" in capsys.readouterr().err


def test_controls_clear(form_doc: Path, tmp_path: Path) -> None:
    out = tmp_path / "cleared.docx"
    # First fill it, then clear it.
    main(["controls", "set", str(form_doc), "--tag", "name", "--value", "Ada", "-o", str(out)])
    assert main(["controls", "clear", str(out), "--tag", "name", "--in-place"]) == 0
    from docx_plus.controls import read_controls

    assert read_controls(Document(str(out)))["name"].is_placeholder is True


def test_controls_set_requires_output(form_doc: Path, capsys: pytest.CaptureFixture[str]) -> None:
    code = main(["controls", "set", str(form_doc), "--tag", "name", "--value", "Ada"])
    assert code == 1
    assert "specify -o/--output or --in-place" in capsys.readouterr().err


# --------------------------------------------------------------------------
# top-level dispatch
# --------------------------------------------------------------------------


def test_no_command_prints_help(capsys: pytest.CaptureFixture[str]) -> None:
    assert main([]) == 2
    assert "usage: docx-plus" in capsys.readouterr().out


def test_version(capsys: pytest.CaptureFixture[str]) -> None:
    with pytest.raises(SystemExit) as exc:
        main(["--version"])
    assert exc.value.code == 0
    assert "docx-plus" in capsys.readouterr().out


def test_unknown_command_exits_2() -> None:
    with pytest.raises(SystemExit) as exc:
        main(["frobnicate"])
    assert exc.value.code == 2
