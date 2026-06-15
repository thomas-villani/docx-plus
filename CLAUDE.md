# CLAUDE.md

Guidance for AI coding agents (and humans) working in this repository.

## What this is

`docx_plus` is an OOXML-level extension layer for [python-docx](https://github.com/python-openxml/python-docx).
It reaches the parts of the `.docx` format python-docx does not expose — the
style cascade, content controls, fields, anchored comments, layout, bookmarks,
footnotes/endnotes, publishing (TOC/captions), tracked changes, and document
protection — while leaving the underlying `Document` object fully usable.

- **Scope discipline:** keep this a lean python-docx *extension*. It is not a
  document-authoring framework and does not do live Word automation. Adjacent
  ideas belong in sibling projects, not here.
- Authoritative API contract: `SPEC.md` (original design) + `ROADMAP.md` (live
  shipped/deferred status). `CHANGELOG.md` is the per-release record.

## Environment & tooling

This project uses **`uv`** for everything. Never call bare `python` or `pip`.

```bash
uv sync --extra dev          # install package + dev deps (single source: pyproject [project.optional-dependencies] dev)
uv run pytest                # run the test suite (configured in [tool.pytest.ini_options])
uv run pytest tests/test_foo.py -k name   # one file / one test
uv run mypy                  # strict type-check (files = ["docx_plus"])
uv run ruff check            # lint  (rules: E,F,W,I,B,UP,D — Google docstrings)
uv run ruff format           # format (line-length 100)
uv run mkdocs serve          # preview docs locally
uv run mkdocs build --strict # docs must build link-clean
```

Pre-commit mirrors the CI lint gate: `uv run pre-commit run --all-files`.

Run an example: `uv run python -m docx_plus.examples.<name>` (e.g. `track_changes`).
Run the CLI: `uv run docx-plus inspect FILE` or `uv run python -m docx_plus.cli`.

## Architecture

Layered, one-way dependencies:

- `core/` — foundation: `DocxPlusError` (base of every typed error), namespace
  map (`ns`), OOXML element helpers (`oxml`), id allocation (`ids`), separate
  OOXML parts (`parts`). Depends on nothing above it.
- **Capability modules** — `styles/`, `controls/`, `fields/`, `comments/`,
  `layout/`, `bookmarks/`, `notes/`, `publishing/`, `revisions/`, `protection/`.
  Each builds on `core/` and is largely independent of its siblings.
- `cli/` — argparse console entry point (`docx-plus`) that composes the
  capability modules. This is the one layer that legitimately imports across
  capabilities.
- `examples/`, `_testing/` — runnable examples and test-only OOXML assertions;
  excluded from coverage and the public API.

Each subpackage's `__init__.py` `__all__` is the authoritative public surface
for that module. `docs/ARCHITECTURE.md` has the full module-by-module breakdown.

## Conventions

- **Errors:** every public exception subclasses `core.DocxPlusError`; dual-inherit
  a stdlib type where it aids `except` ergonomics (e.g.
  `RevisionNotFoundError(DocxPlusError, KeyError)`).
- **Typing:** mypy `strict = True` must pass with zero ignores. `warn_unused_ignores`
  is on. Public APIs are fully typed (`Typing :: Typed`).
- **Docstrings:** Google convention (ruff `D`). Tests, `_testing/`, and
  `examples/` are exempt from `D`.
- **Python:** target 3.10+ (`target-version = py310`); CI tests 3.10–3.13.
- **Coverage:** `fail_under = 90`. New code needs tests.

## Gotchas

- **xpath:** `BaseOxmlElement.xpath()` does NOT accept a `namespaces=` kwarg.
  Use `etree.XPath(expr, namespaces=NSMAP)(node)` instead.
- **Examples must be cp1252-safe:** print ASCII only to stdout so
  `python -m docx_plus.examples.<name>` runs on a default Windows console.
- **Paragraph index base differs by surface:** the CLI `inspect` command numbers
  paragraphs 1-based; the library `read_*` functions use 0-based `paragraph_index`.

## Releasing

```bash
uv run bump-my-version bump {major|minor|patch} --dry-run -v   # preview first; tree must be clean
```

Bumps `pyproject.toml` + `docx_plus/__init__.py`, commits, and tags `vX.Y.Z`.
`CHANGELOG.md` is maintained by hand. After a release, re-stamp the prose docs
(README, `docs/index.md`, `docs/API.md`, `docs/ARCHITECTURE.md`, `docs/SKILLS.md`)
for the new version — these have historically lagged behind the bump.
