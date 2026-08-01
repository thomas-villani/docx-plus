# `docx_plus.publishing.captions`

Figure / table captions: a literal label run (`"Figure "`,
`"Table "`, etc.) followed by a `SEQ` complex field that
auto-numbers items sharing the same caption type. The caption type
is the same name a downstream Table of Figures uses to find
captions (see
[`docx_plus.publishing.figures`](publishing-figures.md)).

## Making a caption referenceable

A `REF` field **cannot point at a `SEQ` field** — only at a bookmark. So
"see Figure 3" is not expressible against a bare caption: there is
nothing for the reference to target. `bookmark_name` (v0.5) closes that
gap by bracketing the label and number in a bookmark.

```python
from docx_plus.bookmarks import add_cross_reference
from docx_plus.publishing import add_caption

cap = doc.add_paragraph()
add_caption(cap, caption_type="Figure", bookmark_name="fig_arch")
cap.add_run(": Architecture overview")      # stays outside the bookmark

body = doc.add_paragraph("As shown in ")
add_cross_reference(body, bookmark="fig_arch")     # -> "Figure 1"
```

The bookmark spans exactly the label run plus the `SEQ` field — the same
extent as Word's own "Only label and number" option — so the reference
resolves to `Figure 1`, not to the description. Anything added with
`add_run` *after* the call falls outside it.

For an anchor the reader never sees, mint a hidden name with
[`BookmarkNameRegistry.next_ref_name`](bookmarks-registry.md).

Architecture walkthrough: [Publishing](../concepts/publishing.md).

::: docx_plus.publishing.captions
    options:
      members:
        - add_caption
