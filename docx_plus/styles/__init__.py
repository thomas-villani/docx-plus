"""Style inspection and modification."""

from docx_plus.styles.inspect import (
    FormattingSource,
    MissingPartError,
    ResolvedFormatting,
    StyleCascadeError,
    resolve_effective_formatting,
)
from docx_plus.styles.modify import (
    InvalidColorError,
    StyleExistsError,
    StyleInfo,
    StyleInUseError,
    StyleNotFoundError,
    StyleProxy,
    UnknownStylePropertyError,
    apply_style,
    create_style,
    delete_style,
    ensure_style,
    find_matching_style,
    list_styles,
    modify_style,
    remap_styles,
)

__all__ = [
    "FormattingSource",
    "InvalidColorError",
    "MissingPartError",
    "ResolvedFormatting",
    "StyleCascadeError",
    "StyleExistsError",
    "StyleInUseError",
    "StyleInfo",
    "StyleNotFoundError",
    "StyleProxy",
    "UnknownStylePropertyError",
    "apply_style",
    "create_style",
    "delete_style",
    "ensure_style",
    "find_matching_style",
    "list_styles",
    "modify_style",
    "remap_styles",
    "resolve_effective_formatting",
]
