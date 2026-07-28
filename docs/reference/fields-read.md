# `docx_plus.fields.read`

Reading the fields a document already contains.

A complex field is not an element — it is a *run sequence* delimited by
`w:fldChar` markers, with the instruction spread across however many
`w:instrText` elements Word happened to split it into. Reading one back
means walking that sequence, which is why this is a capability rather than
an xpath at each call site.

The **instruction** is where a field's meaning lives. The cached `result` is
whatever Word last rendered and can be arbitrarily stale — which is exactly
how a `REF` to a deleted bookmark survives a dozen revisions still showing
the text that used to be correct.

::: docx_plus.fields.read
    options:
      members:
        - read_fields
        - FieldInfo
