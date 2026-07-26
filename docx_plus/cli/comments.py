"""``docx-plus comments`` — list comment threads, resolve or reopen them.

Wraps :func:`docx_plus.comments.read_threads`,
:func:`~docx_plus.comments.resolve_comment`, and
:func:`~docx_plus.comments.reopen_comment`. ``list`` is read-only;
``resolve`` and ``reopen`` mutate and therefore require ``-o/--output``
(or ``--in-place``).

Resolution is thread-wide in Word, and this command inherits that: naming
any comment in a thread resolves or reopens the whole thread.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx_plus.cli._io import (
    dump_json,
    load_document,
    resolve_output,
    save_document,
)
from docx_plus.comments import read_threads, reopen_comment, resolve_comment

if TYPE_CHECKING:
    import argparse

    from docx_plus.comments import AnchoredComment, CommentThread


def register(subparsers: argparse._SubParsersAction[argparse.ArgumentParser]) -> None:
    """Add the ``comments`` subparser and its list/resolve/reopen sub-actions."""
    parser = subparsers.add_parser(
        "comments",
        help="list comment threads, or resolve / reopen them",
        description="Inspect and triage threaded comments.",
    )
    actions = parser.add_subparsers(dest="action", metavar="{list,resolve,reopen}")
    actions.required = True

    list_p = actions.add_parser("list", help="list every comment thread")
    list_p.add_argument("file", help="path to the .docx file")
    list_p.add_argument(
        "--unresolved",
        action="store_true",
        help="show only threads that are still open",
    )
    list_p.add_argument(
        "--json",
        dest="as_json",
        action="store_true",
        help="emit structured JSON instead of text",
    )
    list_p.set_defaults(func=cmd_list)

    resolve_p = actions.add_parser("resolve", help="mark a comment's thread resolved")
    resolve_p.add_argument("file", help="path to the source .docx file")
    resolve_p.add_argument("id", type=int, help="w:id of any comment in the thread")
    _add_output_args(resolve_p)
    resolve_p.set_defaults(func=cmd_resolve)

    reopen_p = actions.add_parser("reopen", help="mark a comment's thread unresolved")
    reopen_p.add_argument("file", help="path to the source .docx file")
    reopen_p.add_argument("id", type=int, help="w:id of any comment in the thread")
    _add_output_args(reopen_p)
    reopen_p.set_defaults(func=cmd_reopen)


def _add_output_args(parser: argparse.ArgumentParser) -> None:
    """Add the shared -o/--output and --in-place options to a mutating action."""
    parser.add_argument("-o", "--output", default=None, help="path to write the result")
    parser.add_argument(
        "--in-place",
        action="store_true",
        help="overwrite the input file instead of requiring -o/--output",
    )


def _comment_json(comment: AnchoredComment) -> dict[str, object]:
    """Render one comment as a JSON-friendly mapping."""
    return {
        "id": comment.comment_id,
        "author": comment.author,
        "initials": comment.initials,
        "timestamp": comment.timestamp,
        "text": comment.text,
        "anchored_text": comment.anchored_text,
        "paragraph_index": comment.paragraph_index,
    }


def _thread_json(thread: CommentThread) -> dict[str, object]:
    """Render one thread as a JSON-friendly mapping."""
    return {
        **_comment_json(thread.root),
        "resolved": thread.resolved,
        "replies": [_comment_json(reply) for reply in thread.replies],
    }


def _print_comment(comment: AnchoredComment, *, indent: str, suffix: str = "") -> None:
    """Print one comment as an indented text line plus its anchor."""
    author = comment.author or "(no author)"
    print(f"{indent}[{comment.comment_id}] {author}{suffix}: {comment.text}")
    if comment.paragraph_index >= 0:
        print(f"{indent}    on paragraph {comment.paragraph_index}: {comment.anchored_text!r}")
    else:
        print(f"{indent}    (orphaned - no anchor in the document body)")


def cmd_list(args: argparse.Namespace) -> int:
    """Handle ``docx-plus comments list``."""
    doc = load_document(args.file)
    threads = read_threads(doc)
    if args.unresolved:
        threads = [thread for thread in threads if not thread.resolved]

    if args.as_json:
        dump_json([_thread_json(thread) for thread in threads])
        return 0

    if not threads:
        print("(no comments)")
        return 0
    for thread in threads:
        _print_comment(
            thread.root,
            indent="",
            suffix=" [resolved]" if thread.resolved else "",
        )
        for reply in thread.replies:
            _print_comment(reply, indent="  ")
    return 0


def cmd_resolve(args: argparse.Namespace) -> int:
    """Handle ``docx-plus comments resolve``."""
    out_path = resolve_output(args)
    doc = load_document(args.file)
    resolve_comment(doc, args.id)
    save_document(doc, out_path)
    print(f"resolved the thread containing comment {args.id}; wrote {out_path}")
    return 0


def cmd_reopen(args: argparse.Namespace) -> int:
    """Handle ``docx-plus comments reopen``."""
    out_path = resolve_output(args)
    doc = load_document(args.file)
    reopen_comment(doc, args.id)
    save_document(doc, out_path)
    print(f"reopened the thread containing comment {args.id}; wrote {out_path}")
    return 0
