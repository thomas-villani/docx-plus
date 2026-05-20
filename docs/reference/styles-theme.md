# `docx_plus.styles.theme`

Read-only theme color **and font** resolution. The cascade resolver in
[`styles.inspect`](styles-inspect.md) uses these to translate `themeColor`
attributes into concrete hex values and `*Theme` font tokens (`minorHAnsi`,
…) into typeface names. Writing themes is out of scope — this module is
read-only (SPEC §1).

::: docx_plus.styles.theme
    options:
      members:
        - load_theme
        - resolve_theme_color
        - resolve_theme_font
        - ThemeColors
        - apply_theme_tint
        - apply_theme_shade
        - apply_lum_mod
        - apply_lum_off
        - ThemeError
