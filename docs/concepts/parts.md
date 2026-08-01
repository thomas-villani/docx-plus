# Separate OOXML parts

v0.1 capabilities (styles, fields, controls, protection) only mutated
the main document part and `settings.xml`. v0.2 introduces three
capabilities backed by **separate** parts that may not exist in a
fresh document:

- `/word/comments.xml` (relationship `RT.COMMENTS`)
- `/word/footnotes.xml` (relationship `RT.FOOTNOTES`)
- `/word/endnotes.xml` (relationship `RT.ENDNOTES`)

v0.4 adds a fourth for comment threading:

- `/word/commentsExtended.xml` (relationship `RT_COMMENTS_EXTENDED`)

v0.5 adds three more — the two remaining comment side-parts, and
numbering:

- `/word/commentsIds.xml` (relationship `RT_COMMENTS_IDS`)
- `/word/people.xml` (relationship `RT_PEOPLE`)
- `/word/numbering.xml` (relationship `RT.NUMBERING`)

`core/parts.py:get_or_create_part(doc, spec)` is the single entry
point. Given a `PartSpec` describing the target, it tries
`doc.part.part_related_by(spec.relationship_type)`; on `KeyError` it
parses `spec.root_xml` for the empty default root element, looks up
the correct part class from `PartFactory.part_type_for`, constructs the
part, and wires the relationship. Returns `(part, root_element)`.

python-docx already registers `CommentsPart` for `WML_COMMENTS` — and
`NumberingPart` for `WML_NUMBERING` — at package-import time. It does
**not** register the footnote, endnote, commentsExtended, commentsIds, or
people content types, so `core/parts.py` does — installing internal
`_FootnotesPart` / `_EndnotesPart` / `_CommentsExtendedPart` /
`_CommentsIdsPart` / `_PeoplePart` subclasses of `XmlPart` with
`PartFactory.part_type_for.setdefault(...)`. Without that registration,
an existing document with footnotes would deserialize the part as the
default `Part` (blob-only), and `part.element` would not exist.

Seven pre-baked `PartSpec` constants cover every need through v0.5:
`COMMENTS_SPEC`, `COMMENTS_EXTENDED_SPEC`, `COMMENTS_IDS_SPEC`,
`PEOPLE_SPEC`, `NUMBERING_SPEC`, `FOOTNOTES_SPEC`, `ENDNOTES_SPEC`.
Custom callers can build their own. The comment side-part content and
relationship types are Microsoft extensions with no member in
python-docx's enums, so they ship alongside as the
`CT_COMMENTS_EXTENDED` / `RT_COMMENTS_EXTENDED`, `CT_COMMENTS_IDS` /
`RT_COMMENTS_IDS`, and `CT_PEOPLE` / `RT_PEOPLE` string constants.

## Numbering is the one that routes around a *broken* path

**`NUMBERING_SPEC` exists because python-docx's own path is broken, not
merely absent.** `DocumentPart.numbering_part` is documented as creating
an empty part when none is present, but it does that through
`NumberingPart.new()` — an unimplemented stub that raises
`NotImplementedError`. It therefore works only for documents that
already carry `numbering.xml`, which the bundled template does, hiding
the failure. Any document from LibreOffice, Pandoc, or a stripped
template hits the stub.

`NUMBERING_SPEC` routes around it, and the [cascade
resolver](cascade.md)'s `_numbering_root` (layer 4) reads the
relationship directly for the same reason — note that `getattr(part,
"numbering_part", None)` does *not* protect you here, since `getattr`'s
default only swallows `AttributeError`.
