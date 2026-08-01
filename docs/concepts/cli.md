# The `docx-plus` CLI

`cli/` is the console entry point registered in `pyproject.toml` as
`docx-plus = "docx_plus.cli:main"` (and runnable as
`python -m docx_plus.cli`). It is a thin argparse shell over the
library: `build_parser()` registers one subparser per subcommand, and
`main(argv)` dispatches to the matching handler, returning `0` on
success, `1` for a handled `DocxPlusError` (printed to stderr), and `2`
when no command is given.

The full command reference — every flag, with worked examples — is
[the CLI page](../cli.md).

Seven subcommands, six of them wrapping one tested library function
each:

- `inspect` — dump effective per-paragraph formatting
  (`styles.resolve_effective_formatting`).
- `restyle` — remap styles onto canonical ids (`styles.remap_styles`).
- `controls` — list / set / clear content-control values
  (`controls`).
- `comments` — list / resolve / reopen comment threads (`comments`).
- `lint` — report formatting defects (`lint.lint`). v0.6.
- `plan` — describe the repair without applying it (`lint.plan_fixes`).
  v0.6.
- `skill` — locate, read, or install the packaged agent skill (v0.5).

Read commands take `--json`; mutating commands require `-o/--output`
(or an explicit `--in-place`) so the input is never overwritten by
accident. Shared load/save plumbing and the `CliError` type live in
`cli/_io.py`, along with the `--rule` / `--exclude` / `--no-tables` /
`--profile` options `lint` and `plan` share.

`lint` and `plan` overload exit `1` as "I found something", which is what
lets either gate a CI step directly. Both are read-only, so there is
nothing to overwrite and the mutating-command convention does not apply
to them.

The CLI is the **one** layer that legitimately imports across
capabilities — it composes `styles/` and `controls/` by design — and is
the documented exception to the [no-cross-imports
invariant](invariants.md#the-invariants).

## `skill` — the packaged agent skill (v0.5)

The LLM-facing guide lived at repo-level `skills/docx-plus/` through
v0.4, which meant `docs/SKILLS.md` claimed the library "ships" it while
linking only to GitHub blob URLs — broken for anyone who had
`pip install`ed. v0.5 moved the tree to `docx_plus/skill/`.

**That move needed no build configuration at all.** Hatchling's
`packages = ["docx_plus"]` already sweeps non-`.py` files — the reason
`py.typed` ships — and the sdist `include` already lists `docx_plus/`.
Verified by building a wheel and unzipping it: all eleven Markdown files
present, then installed into a clean venv with no source tree and
driven through the CLI from there.

`cli/skill.py` is the one command that neither reads nor writes a
`.docx`, so the `-o/--output` / `--in-place` convention does not apply;
it takes `--dest` / `--user` plus a `--force` overwrite guard instead.

Two implementation notes:

- Resources are addressed as `files("docx_plus") / "skill"`, **not**
  `files("docx_plus.skill")`. The latter resolves through the
  namespace-package machinery and yields a `MultiplexedPath`, whose
  `str()` is `MultiplexedPath('…')` — useless as the output of
  `skill path`. Anchoring on the real package and navigating in gives a
  plain `Path`.
- `_copy_tree` walks the `Traversable` API rather than calling
  `shutil.copytree`, so `skill install` works even from a zipimported
  distribution where there is no source directory to copy from.
  `skill path` is the only action that cannot serve that case, and it
  raises a `CliError` saying so.

The suite asserts that every reference topic is linked from `SKILL.md`
and that the frontmatter carries the `name:` / `description:` an agent
needs to discover the skill, so a new topic page cannot land orphaned.

See [the agent skill page](../SKILLS.md) for how to use it.
