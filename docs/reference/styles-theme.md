# `docx_plus.styles.theme`

Read-only theme color resolution. The cascade resolver in
[`styles.inspect`](styles-inspect.md) uses these to translate `themeColor`
attributes into concrete hex values. Writing themes is a v0.2 goal.

::: docx_plus.styles.theme
    options:
      members:
        - load_theme
        - resolve_theme_color
        - ThemeColors
        - apply_theme_tint
        - apply_theme_shade
        - apply_lum_mod
        - apply_lum_off
        - ThemeError
