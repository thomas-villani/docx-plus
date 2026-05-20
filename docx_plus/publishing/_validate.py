"""Shared input validation for publishing helpers.

The TOC / SEQ / Table-of-Figures helpers all interpolate caller-supplied
strings into OOXML complex-field instructions. Unsanitised input lets a
caller's bad value terminate an identifier or quoted argument and inject
arbitrary additional switches, producing files Word silently "repairs"
or that index the wrong things (issues.md H11, M16).

These helpers fail fast at the public surface with `ValueError` per the
SPEC §16 raw-exception carve-out.
"""

from __future__ import annotations

import re

# ECMA-376 17.16.5.56: SEQ identifiers are unquoted single tokens.
# Word accepts ASCII letters / digits / underscore, must start with a
# letter or underscore. Use the same rule as Python identifiers minus
# the unicode allowance.
_SEQ_IDENTIFIER_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")

# ECMA-376 17.16.4.1 general-format picture tokens (the value space for
# the `\*` switch). Word's `Insert → Field` UI emits the capitalised
# forms; the lowercase variants (Roman, alphabetic) are also accepted.
_NUMBERING_PICTURES: frozenset[str] = frozenset(
    {
        "ARABIC",
        "Arabic",
        "ROMAN",
        "Roman",
        "roman",
        "ALPHABETIC",
        "Alphabetic",
        "alphabetic",
        "CardText",
        "DollarText",
        "Hex",
        "HEX",
        "Ordinal",
        "OrdText",
        "MERGEFORMAT",
    }
)


def validate_seq_identifier(name: str, *, arg_name: str) -> str:
    """Return ``name`` if it's a valid SEQ identifier, else raise ValueError.

    Args:
        name: Candidate identifier.
        arg_name: The keyword name being validated, used in the error
            message (e.g. ``"caption_type"``).
    """
    if not name or not _SEQ_IDENTIFIER_RE.fullmatch(name):
        raise ValueError(
            f"{arg_name} must match SEQ identifier rule "
            f"(ASCII letter/underscore start, then letters/digits/underscores); "
            f"got {name!r}"
        )
    return name


def validate_numbering_picture(picture: str) -> str:
    r"""Return ``picture`` if it's a recognised format token, else raise.

    Word silently drops ``\*`` switches with unknown tokens, so callers
    deserve to know.
    """
    if not picture or picture not in _NUMBERING_PICTURES:
        valid = ", ".join(sorted(_NUMBERING_PICTURES))
        raise ValueError(
            f"numbering must be one of {{{valid}}}; got {picture!r}"
        )
    return picture


def validate_outline_levels(levels: tuple[int, int]) -> tuple[int, int]:
    """Validate ``(lo, hi)`` as an outline-level range for a TOC.

    Word's outline levels run 1..9 (the nine heading-level slots).
    Raises ValueError for non-tuples, wrong arity, non-ints, out-of-range
    values, or a reversed range.
    """
    if (
        not isinstance(levels, tuple)
        or len(levels) != 2
        or not all(isinstance(v, int) and not isinstance(v, bool) for v in levels)
    ):
        raise ValueError(
            f"levels must be a tuple of two ints (lo, hi); got {levels!r}"
        )
    lo, hi = levels
    if not (1 <= lo <= hi <= 9):
        raise ValueError(
            f"levels must satisfy 1 <= lo <= hi <= 9 (Word's outline range); "
            f"got ({lo}, {hi})"
        )
    return (lo, hi)


def validate_additional_styles(
    pairs: object,
) -> tuple[tuple[str, int], ...]:
    r"""Validate an ``additional_styles`` iterable for ``add_toc``.

    Returns the canonicalised tuple. Each pair is ``(style_name, level)``
    where ``style_name`` is non-empty and contains no comma or double-quote
    (which would terminate the ``\t`` switch), and ``level`` is in 1..9.
    """
    if pairs is None:
        return ()
    if isinstance(pairs, str) or not hasattr(pairs, "__iter__"):
        raise ValueError(
            f"additional_styles must be an iterable of (str, int) pairs; got {pairs!r}"
        )
    seq = list(pairs)

    out: list[tuple[str, int]] = []
    for i, item in enumerate(seq):
        if (
            not isinstance(item, tuple)
            or len(item) != 2
            or not isinstance(item[0], str)
            or not isinstance(item[1], int)
            or isinstance(item[1], bool)
        ):
            raise ValueError(
                f"additional_styles[{i}] must be a (str, int) tuple; got {item!r}"
            )
        style_name, level = item
        if not style_name or '"' in style_name or "," in style_name:
            raise ValueError(
                f"additional_styles[{i}] style name {style_name!r} must be "
                f"non-empty and contain no comma or double-quote"
            )
        if not (1 <= level <= 9):
            raise ValueError(
                f"additional_styles[{i}] level must be in 1..9; got {level}"
            )
        out.append((style_name, level))
    return tuple(out)


__all__ = [
    "validate_additional_styles",
    "validate_numbering_picture",
    "validate_outline_levels",
    "validate_seq_identifier",
]
