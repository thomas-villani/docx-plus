"""Shared pytest configuration and fixture-provider hooks.

Fixture ``.docx`` files are *generated*, never committed (SPEC §10). This
module is the single canonical generation path: each ``*_docx_path`` fixture
builds only the file it needs, once per session, into a session tmp dir. The
``tests/fixtures/build_fixtures.py`` module's ``main()`` is a separate manual
inspection helper — it does not feed the tests and does not write into the
committed source tree.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.fixtures.build_fixtures import (
    build_empty,
    build_existing_form,
    build_multistyle,
    build_numbered,
    build_themed,
)


@pytest.fixture(scope="session")
def _fixtures_dir(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """Session-scoped tmp dir that holds generated fixtures."""
    return tmp_path_factory.mktemp("docx_plus_fixtures")


@pytest.fixture(scope="session")
def empty_docx_path(_fixtures_dir: Path) -> Path:
    """Path to the freshly-built ``empty.docx`` fixture."""
    return build_empty(_fixtures_dir / "empty.docx")


@pytest.fixture(scope="session")
def multistyle_docx_path(_fixtures_dir: Path) -> Path:
    """Path to the freshly-built ``multistyle.docx`` fixture."""
    return build_multistyle(_fixtures_dir / "multistyle.docx")


@pytest.fixture(scope="session")
def themed_docx_path(_fixtures_dir: Path) -> Path:
    """Path to the freshly-built ``themed.docx`` fixture."""
    return build_themed(_fixtures_dir / "themed.docx")


@pytest.fixture(scope="session")
def existing_form_docx_path(_fixtures_dir: Path) -> Path:
    """Path to the freshly-built ``existing_form.docx`` fixture."""
    return build_existing_form(_fixtures_dir / "existing_form.docx")


@pytest.fixture(scope="session")
def numbered_docx_path(_fixtures_dir: Path) -> Path:
    """Path to the freshly-built ``numbered.docx`` fixture."""
    return build_numbered(_fixtures_dir / "numbered.docx")
