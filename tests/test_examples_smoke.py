"""Smoke-test every example script with no arguments.

Per `IMPLEMENTATION.md` §8: each example runs as part of the test suite so a
public-API change that breaks an example is caught immediately. Subprocess
invocation (rather than importing main()) is deliberate — it exercises the
``python -m docx_plus.examples.<name>`` invocation form documented in each
example's docstring and verifies exit codes the way a user would see them.
"""

from __future__ import annotations

import pkgutil
import subprocess
import sys
from pathlib import Path

import pytest
from docx import Document

import docx_plus.examples

# Discovered from the examples package so a newly-added example is smoke-tested
# automatically (N7). Underscore-prefixed modules (helpers) are excluded.
EXAMPLES = sorted(
    f"docx_plus.examples.{info.name}"
    for info in pkgutil.iter_modules(docx_plus.examples.__path__)
    if not info.name.startswith("_")
)

# Examples that write a .docx into cwd when run with no args. Tuple of
# (module, expected output filename). inspect_document prints only.
WRITES_DOCX = [
    ("docx_plus.examples.restyle_existing", "restyled.docx"),
    ("docx_plus.examples.build_form", "form.docx"),
    ("docx_plus.examples.populate_form", "filled.docx"),
    ("docx_plus.examples.add_comments", "commented.docx"),
    ("docx_plus.examples.multi_column_layout", "multicol.docx"),
    ("docx_plus.examples.bookmarks_and_xrefs", "bookmarks.docx"),
    ("docx_plus.examples.footnotes_and_endnotes", "notes.docx"),
    ("docx_plus.examples.publishing_layout", "publishing.docx"),
]


def _run(module: str, cwd: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, "-m", module],
        cwd=str(cwd),
        capture_output=True,
        text=True,
        timeout=120,
        check=False,
    )


@pytest.mark.parametrize("module", EXAMPLES)
def test_example_runs_with_no_args(module: str, tmp_path: Path) -> None:
    """Every example should exit 0 when invoked with no arguments."""
    result = _run(module, tmp_path)
    assert result.returncode == 0, (
        f"{module} failed with exit {result.returncode}\n"
        f"--- stdout ---\n{result.stdout}\n--- stderr ---\n{result.stderr}"
    )


@pytest.mark.parametrize(("module", "filename"), WRITES_DOCX)
def test_example_output_reopens(module: str, filename: str, tmp_path: Path) -> None:
    """Each docx-writing example must produce a file python-docx can reopen."""
    result = _run(module, tmp_path)
    assert result.returncode == 0, result.stderr
    out_path = tmp_path / filename
    assert out_path.exists(), (
        f"{module} did not write {filename} into cwd. stdout:\n{result.stdout}"
    )
    doc = Document(str(out_path))
    assert doc.paragraphs, f"{filename} from {module} has no paragraphs"
