"""Generate the .docx fixtures used by the test suite.

Fixtures are generated, not committed (SPEC §10). Run as a script to (re)build
all fixtures into ``tests/fixtures/``; the test ``conftest.py`` also calls
:func:`build_all` lazily on first request.

Each builder function is small and explicit so the fixture's contents are
obvious from reading this file.
"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document

from docx_plus.core.oxml import sub

FIXTURES_DIR = Path(__file__).resolve().parent


def build_empty(path: Path) -> Path:
    """Write a minimal valid ``.docx`` with no body content."""
    doc = Document()
    doc.save(path)
    return path


def build_multistyle(path: Path) -> Path:
    """Write a doc with a three-level paragraph style chain Base -> Mid -> Top.

    The chain is intentionally shallow so toggle and basedOn behavior can be
    verified end-to-end by the Phase 2 cascade resolver. One paragraph is
    styled ``Top`` so the cascade has a real target.
    """
    doc = Document()
    styles_element = doc.styles.element

    base = sub(
        styles_element,
        "w:style",
        **{"w:type": "paragraph", "w:styleId": "Base"},
    )
    sub(base, "w:name", **{"w:val": "Base"})
    base_rpr = sub(base, "w:rPr")
    sub(base_rpr, "w:b")  # toggle: bold ON at Base

    mid = sub(
        styles_element,
        "w:style",
        **{"w:type": "paragraph", "w:styleId": "Mid"},
    )
    sub(mid, "w:name", **{"w:val": "Mid"})
    sub(mid, "w:basedOn", **{"w:val": "Base"})
    mid_rpr = sub(mid, "w:rPr")
    sub(mid_rpr, "w:i")  # italic ON at Mid

    top = sub(
        styles_element,
        "w:style",
        **{"w:type": "paragraph", "w:styleId": "Top"},
    )
    sub(top, "w:name", **{"w:val": "Top"})
    sub(top, "w:basedOn", **{"w:val": "Mid"})
    top_rpr = sub(top, "w:rPr")
    sub(top_rpr, "w:b")  # toggle: bold XOR through chain -> effective FALSE

    para = doc.add_paragraph("Styled by Top, which is basedOn Mid, basedOn Base.")
    para_pr = para._p.get_or_add_pPr()
    sub(para_pr, "w:pStyle", **{"w:val": "Top"})

    doc.save(path)
    return path


def build_themed(path: Path) -> Path:
    """Write a doc whose paragraph style uses a themeColor + themeShade.

    Exercises the Phase 2 theme-resolution path end-to-end: the style
    ``Themed`` sets ``w:color`` with ``themeColor="accent1"`` and a
    ``themeShade`` byte, so the resolver has to fetch ``accent1`` from the
    document's theme part and darken it.
    """
    doc = Document()
    styles_element = doc.styles.element

    themed_style = sub(
        styles_element,
        "w:style",
        **{"w:type": "paragraph", "w:styleId": "Themed"},
    )
    sub(themed_style, "w:name", **{"w:val": "Themed"})
    themed_rpr = sub(themed_style, "w:rPr")
    sub(
        themed_rpr,
        "w:color",
        **{
            "w:val": "auto",
            "w:themeColor": "accent1",
            "w:themeShade": "80",
        },
    )

    para = doc.add_paragraph("Themed paragraph.")
    para_pr = para._p.get_or_add_pPr()
    sub(para_pr, "w:pStyle", **{"w:val": "Themed"})

    doc.save(path)
    return path


def build_existing_form(path: Path) -> Path:
    """Write a doc with three SDTs built by hand, *not* via FormBuilder.

    Used by the read-side tests to verify ``read_controls`` works on
    documents that did not originate from ``docx_plus.controls.builder`` —
    third-party tools, Word itself, or earlier hand-rolled scripts.
    """
    doc = Document()
    para = doc.add_paragraph("Name: ")
    para_two = doc.add_paragraph("Region: ")
    para_three = doc.add_paragraph("Subscribe: ")

    # Filled text SDT (no showingPlcHdr).
    text_sdt = sub(para._p, "w:sdt")
    text_pr = sub(text_sdt, "w:sdtPr")
    sub(text_pr, "w:tag", **{"w:val": "name"})
    sub(text_pr, "w:id", **{"w:val": "100"})
    sub(text_pr, "w:text")
    text_content = sub(text_sdt, "w:sdtContent")
    text_run = sub(text_content, "w:r")
    text_t = sub(text_run, "w:t")
    text_t.text = "Ada Lovelace"

    # Dropdown SDT in placeholder state.
    dd_sdt = sub(para_two._p, "w:sdt")
    dd_pr = sub(dd_sdt, "w:sdtPr")
    sub(dd_pr, "w:alias", **{"w:val": "Region selector"})
    sub(dd_pr, "w:tag", **{"w:val": "region"})
    sub(dd_pr, "w:id", **{"w:val": "200"})
    sub(dd_pr, "w:showingPlcHdr")
    dd_list = sub(dd_pr, "w:dropDownList")
    sub(dd_list, "w:listItem", **{"w:displayText": "Choose a region", "w:value": ""})
    sub(dd_list, "w:listItem", **{"w:displayText": "North", "w:value": "N"})
    sub(dd_list, "w:listItem", **{"w:displayText": "South", "w:value": "S"})
    dd_content = sub(dd_sdt, "w:sdtContent")
    dd_run = sub(dd_content, "w:r")
    dd_run_pr = sub(dd_run, "w:rPr")
    sub(dd_run_pr, "w:rStyle", **{"w:val": "PlaceholderText"})
    dd_t = sub(dd_run, "w:t")
    dd_t.text = "Choose a region"

    # Checkbox SDT, checked.
    cb_sdt = sub(para_three._p, "w:sdt")
    cb_pr = sub(cb_sdt, "w:sdtPr")
    sub(cb_pr, "w:tag", **{"w:val": "subscribe"})
    sub(cb_pr, "w:id", **{"w:val": "300"})
    cb_box = sub(cb_pr, "w14:checkbox")
    sub(cb_box, "w14:checked", **{"w14:val": "1"})
    sub(cb_box, "w14:checkedState", **{"w14:val": "2612", "w14:font": "MS Gothic"})
    sub(cb_box, "w14:uncheckedState", **{"w14:val": "2610", "w14:font": "MS Gothic"})
    cb_content = sub(cb_sdt, "w:sdtContent")
    cb_run = sub(cb_content, "w:r")
    cb_t = sub(cb_run, "w:t")
    cb_t.text = "☒"

    doc.save(path)
    return path


def build_all(out_dir: Path = FIXTURES_DIR) -> dict[str, Path]:
    """Build every Phase 1 fixture into ``out_dir`` and return their paths."""
    out_dir.mkdir(parents=True, exist_ok=True)
    return {
        "empty": build_empty(out_dir / "empty.docx"),
        "multistyle": build_multistyle(out_dir / "multistyle.docx"),
        "themed": build_themed(out_dir / "themed.docx"),
        "existing_form": build_existing_form(out_dir / "existing_form.docx"),
    }


def main() -> int:
    """Entry point for ``python -m tests.fixtures.build_fixtures``."""
    built = build_all()
    for name, path in built.items():
        print(f"built {name}: {path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
