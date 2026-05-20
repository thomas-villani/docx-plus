# `docx_plus.core.ids`

The SDT `w:id` allocator. One registry per document edit session. Other
ID namespaces (`r:id`, bookmark IDs, comment IDs) are separate
uniqueness domains and will get their own registries in later phases.

::: docx_plus.core.ids
    options:
      members:
        - IdRegistry
        - DuplicateIdError
        - IdRangeError
