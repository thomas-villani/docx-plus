# `docx_plus.comments.read`

Inverse of `add_comment`: walks `comments.xml` and pairs each
`<w:comment>` with the body-side range it anchors. Each result carries
the comment body text, the anchored document text, the paragraph index
where the comment is attached, and parsed metadata (author, initials,
timestamp). Orphaned comments (no matching body range) appear with
`anchored_text=""` and `paragraph_index=-1`.

::: docx_plus.comments.read
    options:
      members:
        - read_comments
        - AnchoredComment
