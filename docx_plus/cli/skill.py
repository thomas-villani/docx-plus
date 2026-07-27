"""``docx-plus skill`` — locate, read, or install the packaged agent skill.

The library ships an LLM-facing guide under ``docx_plus/skill/``: a
``SKILL.md`` entry point plus topic reference files an agent pulls in on
demand. It rides along in the wheel because hatchling's
``packages = ["docx_plus"]`` sweeps non-``.py`` files — the same reason
``py.typed`` ships — so ``pip install docx-plus`` puts it on disk with no
extra build configuration.

This command is the seam between that and an agent's skills directory:

- ``path`` — print where the packaged skill lives.
- ``list`` — list the reference topics.
- ``show [TOPIC]`` — print ``SKILL.md``, or one reference file.
- ``install`` — copy the tree into a skills directory.

Unlike every other ``docx-plus`` command, none of these read or write a
``.docx``, so the ``-o/--output`` / ``--in-place`` convention does not
apply. ``install`` takes ``--dest`` / ``--user`` instead, and refuses to
overwrite without ``--force``.
"""

from __future__ import annotations

import shutil
from importlib.resources import files
from pathlib import Path
from typing import TYPE_CHECKING

from docx_plus.cli._io import CliError

if TYPE_CHECKING:
    import argparse
    from importlib.abc import Traversable

#: Directory an agent skill is installed under, relative to a skills root.
SKILL_NAME = "docx-plus"

#: Where Claude Code looks for project-scoped and user-scoped skills.
PROJECT_SKILLS_DIR = Path(".claude") / "skills"
USER_SKILLS_DIR = Path.home() / ".claude" / "skills"


def register(subparsers: argparse._SubParsersAction[argparse.ArgumentParser]) -> None:
    """Add the ``skill`` subparser and its path/list/show/install actions."""
    parser = subparsers.add_parser(
        "skill",
        help="locate, read, or install the packaged agent skill",
        description="Work with the LLM-facing guide bundled in the wheel.",
    )
    actions = parser.add_subparsers(dest="action", metavar="{path,list,show,install}")
    actions.required = True

    path_p = actions.add_parser("path", help="print the packaged skill directory")
    path_p.set_defaults(func=cmd_path)

    list_p = actions.add_parser("list", help="list the reference topics")
    list_p.set_defaults(func=cmd_list)

    show_p = actions.add_parser("show", help="print SKILL.md or one reference topic")
    show_p.add_argument(
        "topic",
        nargs="?",
        default=None,
        help="reference topic (see `docx-plus skill list`); omit for SKILL.md",
    )
    show_p.set_defaults(func=cmd_show)

    install_p = actions.add_parser("install", help="copy the skill into a skills directory")
    install_p.add_argument(
        "--dest",
        default=None,
        help=(
            "skills directory to install into; the skill is written to "
            f"<dest>/{SKILL_NAME}. Defaults to ./{PROJECT_SKILLS_DIR}"
        ),
    )
    install_p.add_argument(
        "--user",
        action="store_true",
        help=f"install into {USER_SKILLS_DIR} instead of the current project",
    )
    install_p.add_argument(
        "--force",
        action="store_true",
        help="overwrite an existing installation",
    )
    install_p.set_defaults(func=cmd_install)


def cmd_path(args: argparse.Namespace) -> int:
    """Print the directory holding the packaged skill."""
    print(_skill_dir())
    return 0


def cmd_list(args: argparse.Namespace) -> int:
    """List the reference topics available to ``show``."""
    for topic in _topics():
        print(topic)
    return 0


def cmd_show(args: argparse.Namespace) -> int:
    """Print ``SKILL.md``, or the named reference topic."""
    root = _skill_root()
    if args.topic is None:
        print(_read(root / "SKILL.md"))
        return 0

    topic = args.topic.removesuffix(".md")
    available = _topics()
    if topic not in available:
        raise CliError(f"unknown topic {args.topic!r}; available: {', '.join(available)}")
    print(_read(root / "reference" / f"{topic}.md"))
    return 0


def cmd_install(args: argparse.Namespace) -> int:
    """Copy the packaged skill tree into a skills directory."""
    if args.user and args.dest is not None:
        raise CliError("--user and --dest are mutually exclusive")

    if args.dest is not None:
        skills_root = Path(args.dest).expanduser()
    elif args.user:
        skills_root = USER_SKILLS_DIR
    else:
        skills_root = Path.cwd() / PROJECT_SKILLS_DIR

    target = skills_root / SKILL_NAME
    if target.exists():
        if not args.force:
            raise CliError(f"{target} already exists; pass --force to overwrite")
        shutil.rmtree(target)

    count = _copy_tree(_skill_root(), target)
    print(f"installed {count} files to {target}")
    return 0


# ---------------------------------------------------------------------------
# Internals.
# ---------------------------------------------------------------------------


def _skill_root() -> Traversable:
    """Return the packaged skill directory as an importlib Traversable.

    Anchored on ``docx_plus`` and navigated into rather than addressed as
    ``docx_plus.skill`` directly: the latter resolves through the
    namespace-package machinery and yields a ``MultiplexedPath``, which
    has no filesystem path to print.
    """
    return files("docx_plus") / "skill"


def _skill_dir() -> Path:
    """Return the packaged skill directory as a real filesystem path.

    Raises:
        CliError: If the distribution is a zipimport (an unextracted
            ``.zip``/``.egg``), where the tree has no on-disk location.
            Installed wheels are always unpacked, so this is a corner
            case — ``skill install`` still works there.
    """
    root = _skill_root()
    try:
        return Path(str(root)).resolve(strict=True)
    except (OSError, TypeError) as exc:
        raise CliError(
            "the packaged skill has no filesystem path (zipped distribution); "
            "use `docx-plus skill install` to extract it"
        ) from exc


def _topics() -> list[str]:
    """Return the reference topic names, without the ``.md`` suffix."""
    reference = _skill_root() / "reference"
    if not reference.is_dir():  # pragma: no cover - packaging would be broken
        return []
    return sorted(
        entry.name.removesuffix(".md")
        for entry in reference.iterdir()
        if entry.name.endswith(".md")
    )


def _read(resource: Traversable) -> str:
    """Read a packaged Markdown file as UTF-8, stripped of its trailing newline."""
    return resource.read_text(encoding="utf-8").rstrip("\n")


def _copy_tree(source: Traversable, target: Path) -> int:
    """Recursively copy a Traversable tree onto the filesystem.

    Walks the resource API rather than using :func:`shutil.copytree` so
    it works whether the distribution is unpacked on disk or read out of
    a zip.

    Returns:
        The number of files written.
    """
    target.mkdir(parents=True, exist_ok=True)
    written = 0
    for entry in source.iterdir():
        destination = target / entry.name
        if entry.is_dir():
            written += _copy_tree(entry, destination)
        else:
            destination.write_bytes(entry.read_bytes())
            written += 1
    return written


__all__ = ["cmd_install", "cmd_list", "cmd_path", "cmd_show", "register"]
