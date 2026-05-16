"""Shared pytest configuration and fixture-provider hooks."""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.fixtures.build_fixtures import build_all

FIXTURES_DIR = Path(__file__).parent / "fixtures"


@pytest.fixture(scope="session")
def fixtures(tmp_path_factory: pytest.TempPathFactory) -> dict[str, Path]:
    """Build every Phase 1 fixture once per session into a tmp dir."""
    out = tmp_path_factory.mktemp("docx_plus_fixtures")
    return build_all(out)


@pytest.fixture
def empty_docx_path(fixtures: dict[str, Path]) -> Path:
    """Path to the freshly-built ``empty.docx`` fixture."""
    return fixtures["empty"]


@pytest.fixture
def multistyle_docx_path(fixtures: dict[str, Path]) -> Path:
    """Path to the freshly-built ``multistyle.docx`` fixture."""
    return fixtures["multistyle"]


@pytest.fixture
def themed_docx_path(fixtures: dict[str, Path]) -> Path:
    """Path to the freshly-built ``themed.docx`` fixture."""
    return fixtures["themed"]


@pytest.fixture
def existing_form_docx_path(fixtures: dict[str, Path]) -> Path:
    """Path to the freshly-built ``existing_form.docx`` fixture."""
    return fixtures["existing_form"]
