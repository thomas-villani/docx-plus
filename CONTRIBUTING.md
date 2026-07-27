# Contributing to docx_plus

Thanks for taking an interest. This document covers the development
setup, the quality gates a change has to pass, and the conventions that
keep the codebase coherent.

By participating you agree to abide by the
[Code of Conduct](CODE_OF_CONDUCT.md).

## Scope

`docx_plus` is an **extension layer for python-docx**, not a
document-authoring framework and not a Word automation tool. Every
feature either fills a documented python-docx gap or rounds out a
surface already started here.

Before opening a large PR, check [`ROADMAP.md`](ROADMAP.md) — it records
what is shipped, what is backlogged, and what has been explicitly
[considered and declined](ROADMAP.md#considered-not-on-the-roadmap).
If your idea is on the declined list, an issue arguing the case is more
useful than a PR.

## Development setup

The project uses [`uv`](https://docs.astral.sh/uv/) for everything.
Never call bare `python` or `pip`.

```bash
git clone https://github.com/thomas-villani/docx-plus.git
cd docx-plus
uv sync --extra dev          # install package + dev dependencies
uv run pre-commit install    # ruff check + ruff format on every commit
```

Dev dependencies have a single source of truth:
`[project.optional-dependencies] dev` in `pyproject.toml`. The older
`[tool.uv] dev-dependencies` table is deliberately not duplicated so the
two cannot drift.

## Quality gates

CI runs all of these. Run them locally before pushing:

```bash
uv run pytest                  # test suite
uv run mypy                    # strict type check, zero ignores
uv run ruff check              # lint  (E,F,W,I,B,UP,D — Google docstrings)
uv run ruff format             # format (line length 100)
uv run mkdocs build --strict   # docs must build link-clean
```

Or in one shot, mirroring the CI lint gate:

```bash
uv run pre-commit run --all-files
```

Targeted test runs:

```bash
uv run pytest tests/test_styles_inspect.py
uv run pytest tests/test_styles_inspect.py -k provenance
```

CI additionally runs a **lower-bound dependency** job
(`python-docx==1.0.0`, `lxml==4.9.0`), tests Python 3.10–3.13 on Linux,
and adds a Windows job on 3.13 — the primary dev box is Windows, and it
catches path handling and lxml CRLF behaviour that Linux-only CI misses.
If your change needs a newer API from either dependency, raise the floor
in `pyproject.toml` and say why in the PR.

Some smoke tests need LibreOffice (`soffice`) on `PATH` and are marked
`requires_libreoffice`; they skip cleanly when it is absent.

## Conventions

**Architecture.** Dependencies are one-way. `core/` (errors, namespaces,
OOXML helpers, id allocation, parts) depends on nothing above it.
Capability modules (`styles/`, `controls/`, `fields/`, `comments/`,
`layout/`, `bookmarks/`, `notes/`, `publishing/`, `tables/`,
`numbering/`, `revisions/`, `protection/`) build on `core/` and are
largely independent of each other. `cli/` is the one layer that
legitimately imports across capabilities. See
[`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) for the full breakdown.

**Public surface.** Each subpackage's `__init__.py` `__all__` is
authoritative. If a symbol is not in `__all__`, it is not public.

**Errors.** Every public exception subclasses `core.DocxPlusError`.
Dual-inherit a stdlib type where it aids `except` ergonomics — e.g.
`RevisionNotFoundError(DocxPlusError, KeyError)`.

**Typing.** `mypy --strict` must pass with zero ignores;
`warn_unused_ignores` is on. Public APIs are fully typed.

**Docstrings.** Google convention, enforced by ruff's `D` rules. Tests,
`_testing/`, and `examples/` are exempt.

**Coverage.** `fail_under = 90`. New code needs tests.

**Examples must be cp1252-safe.** Print ASCII only to stdout so
`uv run python -m docx_plus.examples.<name>` runs on a default Windows
console.

**xpath gotcha.** `BaseOxmlElement.xpath()` does *not* accept a
`namespaces=` keyword. Use
`etree.XPath(expr, namespaces=NSMAP)(node)` instead.

## Verifying against Word

OOXML is a format with a single reference implementation, and guessing
at what Word writes is how subtly-wrong output ships. For anything
touching a part or attribute the project has not written before, the
expectation is to **have Word author the file first**, unzip it, and
match the observed markup — several v0.5 details (hex durable ids, four
content-type URIs) were settled exactly this way.

If you cannot run Word, say so in the PR. A reviewer with a copy can
verify, and it is better to flag the gap than to assert an unverified
shape.

## Pull requests

- Branch off `main`; keep the change focused on one concern.
- Conventional commit subjects (`feat:`, `fix:`, `docs:`, `chore:`,
  `test:`, `refactor:`).
- Add tests. New OOXML output is best asserted with the helpers in
  `docx_plus/_testing/`.
- Update [`CHANGELOG.md`](CHANGELOG.md) under `## [Unreleased]`. The
  changelog is maintained by hand and records *why* a change was made,
  not just what — match the surrounding tone.
- Update `ROADMAP.md` if the change lands or reshapes a roadmap item.
- Docs for user-visible surface: docstrings feed the auto-generated
  reference, but the hand-curated [`docs/API.md`](docs/API.md) index
  needs a matching entry.

## Reporting bugs

Open an issue with the [bug report
form](https://github.com/thomas-villani/docx-plus/issues/new?template=bug_report.yml).
The single most useful thing you can attach is a **minimal `.docx`**
that reproduces the problem, or the offending XML fragment from
`word/document.xml`. Include the `docx_plus`, `python-docx`, `lxml`, and
Python versions, and say which Word version you observed the behaviour
in — the format's edge cases are often version-specific.

## Releasing

Maintainer-only, recorded here so the process is not folklore:

```bash
uv run bump-my-version bump {major|minor|patch} --dry-run -v   # preview; tree must be clean
uv run bump-my-version bump minor
git push --follow-tags
```

The bump updates `pyproject.toml`, `docx_plus/__init__.py`, and the
project's own entry in `uv.lock`, then commits and tags `vX.Y.Z`.
Pushing the tag triggers `.github/workflows/release.yml`, which
re-runs the full matrix as a release gate and publishes to PyPI via
trusted publishing.

`CHANGELOG.md` is maintained by hand. After a release, re-stamp the
prose docs (README, `docs/index.md`, `docs/API.md`,
`docs/ARCHITECTURE.md`, `docs/SKILLS.md`) for the new version — these
have historically lagged behind the bump.

## License

By contributing, you agree that your contributions are licensed under
the [MIT License](LICENSE) that covers the project.
