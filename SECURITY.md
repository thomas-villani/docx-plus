# Security Policy

## Supported versions

`docx_plus` is a small library on a rolling release. Security fixes land
on the latest minor version; there are no long-term support branches.

| Version | Supported |
|---|---|
| 0.5.x | Yes |
| < 0.5 | No — upgrade |

## Reporting a vulnerability

**Do not open a public issue for a security problem.**

Report it privately through GitHub's
[private vulnerability reporting](https://github.com/thomas-villani/docx-plus/security/advisories/new),
or by email to <thomas.villani@gmail.com>.

Please include:

- a description of the issue and its impact,
- a minimal `.docx` or code sample that reproduces it,
- the `docx_plus`, `python-docx`, `lxml`, and Python versions.

You can expect an acknowledgement within a week. This is a small
single-maintainer project, so please size your expectations
accordingly — but a real vulnerability will be prioritised over
feature work.

## Threat model

Worth being explicit about, because the realistic risk here is not in
this library's own code.

`docx_plus` parses and writes OOXML, and it does so **through
[`lxml`](https://lxml.de/) and [`python-docx`](https://github.com/python-openxml/python-docx)**.
A `.docx` file is a ZIP archive of XML. If you process documents from
untrusted sources, the attack surface that matters is XML and ZIP
parsing — entity expansion, external entity resolution, decompression
bombs — and that surface belongs to those upstream dependencies, not to
this package. Keep them current, and treat untrusted documents as
untrusted input regardless of which library opens them.

Within `docx_plus` itself, the things I would consider reportable:

- Any path where library input causes a write outside the intended
  output file — notably `docx-plus skill install`, which copies a tree
  to a user-supplied destination.
- Crafted document content that escapes an OOXML value context and
  injects markup into a neighbouring element (an XML-injection bug in
  how a value is written).
- A CLI command mutating its input file when the caller did not pass
  `--in-place`.

Not vulnerabilities:

- Producing a document Word declines to open. That is a correctness
  bug — please file it as a normal issue, they are welcome.
- Resource exhaustion from a deliberately malformed archive handed
  straight to `python-docx`. Report that upstream.
- `protect_document` being circumventable. Word's document protection is
  an editing convention, **not a security boundary** — it is trivially
  removed by anyone who can open the ZIP, and neither Word nor this
  library claims otherwise. Do not use it to protect secrets.
