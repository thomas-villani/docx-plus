# `docx_plus.numbering.registry`

Allocators for the two disjoint id namespaces in `numbering.xml`.

Both allocate with `next_sequential()` — the lowest free integer — rather
than the random `next()` every other namespace uses. Word and python-docx
both number lists this way, and these ids are read by humans debugging
list behaviour far more often than most.

The two namespaces differ at the bottom of their range, and it matters:
`w:abstractNumId` legitimately starts at **0** (the bundled template uses
0–8), while `w:numId` starts at **1** because `0` inside a `w:numPr` is
the sentinel meaning "no numbering" — the only way a paragraph opts out
of a list applied by its style.

::: docx_plus.numbering.registry
    options:
      members:
        - NumIdRegistry
        - AbstractNumIdRegistry
