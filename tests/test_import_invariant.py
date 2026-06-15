"""Enforce SPEC §9.1: capability modules import from ``core/`` only.

Walks every ``.py`` file under each capability subpackage with ``ast``, scans
``Import`` / ``ImportFrom`` nodes, and asserts no module under one capability
ever imports from another capability. Detection is automatic for capability
directories added later (e.g. ``sections/`` in v0.2).

The AST walk resolves only **absolute** imports. To keep the cross-capability
check sound, a companion test forbids relative imports (``from . import x`` /
``from ..fields import y``) outright — a relative hop into a sibling capability
would otherwise slip past the absolute-name match. One residual gap remains by
design: dynamic imports (``importlib.import_module(...)``) are invisible to a
static AST scan; the library does not use them in capability code, and this is
documented rather than enforced.
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
    "revisions",
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


def _relative_imports(tree: ast.AST) -> list[str]:
    """Return a description of every relative ``ImportFrom`` (``level > 0``)."""
    out: list[str] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom) and node.level > 0:
            dots = "." * node.level
            out.append(f"from {dots}{node.module or ''} import ...")
    return out


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


@pytest.mark.parametrize(("capability", "py_file"), _module_files())
def test_no_relative_imports(capability: str, py_file: Path) -> None:
    """Forbid relative imports so the absolute-name cross-capability check holds.

    A relative hop (``from ..fields import X``) would otherwise be invisible to
    :func:`test_no_cross_capability_imports`, which only resolves absolute
    module names.
    """
    tree = ast.parse(py_file.read_text(encoding="utf-8"))
    relative = _relative_imports(tree)
    assert not relative, (
        f"{py_file.relative_to(PACKAGE_ROOT)} uses relative imports "
        f"(absolute imports only, SPEC §9.1): {relative}"
    )
