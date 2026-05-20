"""Enforce SPEC §9.1: capability modules import from ``core/`` only.

Walks every ``.py`` file under each capability subpackage with ``ast``, scans
``Import`` / ``ImportFrom`` nodes, and asserts no module under one capability
ever imports from another capability. Detection is automatic for capability
directories added later (e.g. ``sections/`` in v0.2).
"""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

PACKAGE_ROOT = Path(__file__).resolve().parent.parent / "docx_plus"

CAPABILITIES = {
    "styles",
    "controls",
    "fields",
    "protection",
    "comments",
    "layout",
    "bookmarks",
    "notes",
    "publishing",
}


def _module_files() -> list[tuple[str, Path]]:
    out: list[tuple[str, Path]] = []
    for capability in CAPABILITIES:
        capability_dir = PACKAGE_ROOT / capability
        if not capability_dir.is_dir():
            continue
        for py_file in capability_dir.rglob("*.py"):
            out.append((capability, py_file))
    return out


def _imported_modules(tree: ast.AST) -> set[str]:
    seen: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                seen.add(alias.name)
        elif isinstance(node, ast.ImportFrom):
            if node.module:
                seen.add(node.module)
    return seen


@pytest.mark.parametrize(("capability", "py_file"), _module_files())
def test_no_cross_capability_imports(capability: str, py_file: Path) -> None:
    forbidden = {f"docx_plus.{other}" for other in CAPABILITIES if other != capability}
    tree = ast.parse(py_file.read_text(encoding="utf-8"))
    imported = _imported_modules(tree)
    violations = {
        imp for imp in imported if any(imp == f or imp.startswith(f + ".") for f in forbidden)
    }
    assert not violations, (
        f"{py_file.relative_to(PACKAGE_ROOT)} imports from another capability: {violations}"
    )
