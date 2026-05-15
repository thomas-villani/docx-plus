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


def build_all(out_dir: Path = FIXTURES_DIR) -> dict[str, Path]:
    """Build every Phase 1 fixture into ``out_dir`` and return their paths."""
    out_dir.mkdir(parents=True, exist_ok=True)
    return {
        "empty": build_empty(out_dir / "empty.docx"),
        "multistyle": build_multistyle(out_dir / "multistyle.docx"),
        "themed": build_themed(out_dir / "themed.docx"),
    }


def main() -> int:
    """Entry point for ``python -m tests.fixtures.build_fixtures``."""
    built = build_all()
    for name, path in built.items():
        print(f"built {name}: {path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
