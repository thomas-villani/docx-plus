# `docx_plus.core.errors`

The library base exception. `DocxPlusError` is the root of every typed
error in docx_plus — each submodule subclasses it so callers can catch the
library's failures without also catching unrelated `ValueError` /
`RuntimeError` instances from python-docx or lxml.

::: docx_plus.core.errors
    options:
      members:
        - DocxPlusError
