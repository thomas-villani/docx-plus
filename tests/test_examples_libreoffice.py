"""LibreOffice headless rendering smoke tests (SPEC §13 Layer 3).

Converts every example's output to PDF via ``soffice --headless --convert-to
pdf`` and asserts: exit 0, a non-empty PDF lands in the output directory.
Catches the broad class of "structurally-invalid .docx" bugs that
``python-docx`` happily re-opens but a real renderer chokes on — especially
relevant for the Phase 4 SDT + ``w14`` namespace plumbing and the Phase 5
``settings.xml`` insertions.

Gated behind ``pytest.mark.requires_libreoffice`` (declared in
``pyproject.toml``) and additionally autoskipped via :func:`shutil.which`
when ``soffice`` is not on PATH, so this suite is a no-op on dev boxes
without LibreOffice installed. The marker exists so CI can run *only* this
layer with ``pytest -m requires_libreoffice`` on a runner that has
``soffice``.
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

import pytest

from tests.test_examples_smoke import WRITES_DOCX

_SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")

pytestmark = [
    pytest.mark.requires_libreoffice,
    pytest.mark.skipif(_SOFFICE is None, reason="soffice / libreoffice not on PATH"),
]


def _run_example(module: str, cwd: Path) -> None:
    """Subprocess-invoke an example with no args; raises on non-zero exit."""
    result = subprocess.run(
        [sys.executable, "-m", module],
        cwd=str(cwd),
        capture_output=True,
        text=True,
        timeout=120,
        check=False,
    )
    if result.returncode != 0:
        raise AssertionError(
            f"{module} failed: exit {result.returncode}\n"
            f"stdout:\n{result.stdout}\nstderr:\n{result.stderr}"
        )


def _soffice_to_pdf(docx_path: Path, out_dir: Path) -> Path:
    """Convert ``docx_path`` to PDF in ``out_dir`` via soffice headless."""
    assert _SOFFICE is not None
    result = subprocess.run(
        [
            _SOFFICE,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(out_dir),
            str(docx_path),
        ],
        capture_output=True,
        text=True,
        timeout=180,
        check=False,
    )
    if result.returncode != 0:
        raise AssertionError(
            f"soffice convert failed: exit {result.returncode}\n"
            f"stdout:\n{result.stdout}\nstderr:\n{result.stderr}"
        )
    pdf_path = out_dir / (docx_path.stem + ".pdf")
    if not pdf_path.exists():
        raise AssertionError(
            f"soffice exited 0 but {pdf_path} not produced. stdout:\n{result.stdout}"
        )
    return pdf_path


# Render every docx-writing example — including the full v0.2 surface
# (comments, layout, bookmarks, notes, publishing) — so a structurally
# invalid file that python-docx tolerates but a real renderer rejects is
# caught here (M22). Shares WRITES_DOCX with the smoke suite to avoid drift.
@pytest.mark.parametrize(("module", "filename"), WRITES_DOCX)
def test_example_renders_to_pdf(module: str, filename: str, tmp_path: Path) -> None:
    """Each example's docx output must convert to a non-empty PDF."""
    _run_example(module, tmp_path)
    docx_path = tmp_path / filename
    assert docx_path.exists(), f"{module} did not produce {filename}"

    pdf_path = _soffice_to_pdf(docx_path, tmp_path)
    assert pdf_path.stat().st_size > 0, f"{pdf_path} is empty"
