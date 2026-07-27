<!--
Thanks for contributing. Keep the change focused on one concern; see
CONTRIBUTING.md for the conventions and quality gates.
-->

## What this changes

<!-- A sentence or two. If it closes an issue, say "Closes #123". -->

## Why

<!--
The reasoning, not just the mechanics. If this fills a python-docx gap,
name it. If it changes existing OOXML output, say what Word did before and
what it does now.
-->

## How it was verified

<!--
Delete what does not apply.
-->

- [ ] New tests cover the change
- [ ] Verified the OOXML against a file **Word itself authored** (see
      CONTRIBUTING.md § Verifying against Word)
- [ ] Round-trips through save / reopen
- [ ] Opened the output in Word / LibreOffice and confirmed it renders

<!-- If you could not verify against Word, say so here — flagging the gap is
     better than asserting an unverified shape. -->

## Checklist

- [ ] `uv run pytest` passes
- [ ] `uv run mypy` passes (strict, zero ignores)
- [ ] `uv run ruff check` and `uv run ruff format` are clean
- [ ] `uv run mkdocs build --strict` builds link-clean
- [ ] `CHANGELOG.md` updated under `## [Unreleased]`
- [ ] Public symbols are exported in the subpackage `__all__` and listed in
      `docs/API.md`
- [ ] `ROADMAP.md` updated, if this lands or reshapes a roadmap item

## Breaking changes

<!-- None, or describe the migration. -->
