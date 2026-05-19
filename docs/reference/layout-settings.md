# `docx_plus.layout.settings`

Doc-level `<w:evenAndOddHeaders/>` switch in `settings.xml`. This flag
is constantly confused with the per-section `titlePg` that
python-docx already exposes via
`Section.different_first_page_header_footer`. The flag here is
different: it tells Word that even-numbered pages may have a different
header/footer from odd-numbered pages, *across every section*.

Schema-strict insertion via `core.insert_before_first_anchor`. Both
toggle functions are idempotent.

::: docx_plus.layout.settings
    options:
      members:
        - enable_distinct_even_odd_headers
        - disable_distinct_even_odd_headers
