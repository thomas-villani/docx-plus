"""Style inspection and modification (Phase 2 inspection live; Phase 3 to come)."""

from docx_plus.styles.inspect import (
    FormattingSource,
    MissingPartError,
    ResolvedFormatting,
    StyleCascadeError,
    resolve_effective_formatting,
)

__all__ = [
    "FormattingSource",
    "MissingPartError",
    "ResolvedFormatting",
    "StyleCascadeError",
    "resolve_effective_formatting",
]
