# `docx_plus.styles.modify`

Style creation, modification, application, removal, and reconciliation.
Field names for `**properties` kwargs match [`ResolvedFormatting`][docx_plus.styles.inspect.ResolvedFormatting]
so cascade output round-trips back through the modifier without
translation.

Schema-strict child ordering for `w:style`, `w:pPr`, and `w:rPr` is
enforced internally — see [`ARCHITECTURE.md` §3](../ARCHITECTURE.md#3-schema-strict-insertion).

The Phase 3.5 remap surface — `find_matching_style`, `remap_styles`, and
`ensure_style(match_existing=True)` — is documented in
[`ARCHITECTURE.md` §4](../ARCHITECTURE.md#4-style-remapping-phase-35).

::: docx_plus.styles.modify
    options:
      members:
        - create_style
        - modify_style
        - apply_style
        - delete_style
        - ensure_style
        - find_matching_style
        - remap_styles
        - list_styles
        - StyleProxy
        - StyleInfo
        - InvalidColorError
        - StyleExistsError
        - StyleNotFoundError
        - StyleInUseError
        - UnknownStylePropertyError
