# Package layout

```
docx_plus/
├── __init__.py              # top-level re-exports (DocxPlusError, __version__)
├── py.typed                 # PEP 561 marker — the package ships its own types
├── core/                    # foundation primitives — every capability depends on these
│   ├── __init__.py          # re-exports the public surface
│   ├── errors.py            # DocxPlusError — the base of every typed error
│   ├── borders.py           # Border, border_attrs — shared by page, table, cell borders
│   ├── ns.py                # W, W14, W15, R, MC, A, XML constants + NSMAP / BUILD_NSMAP + qn()
│   ├── oxml.py              # el(), sub(), xpath(), remove(),
│   │                        # build_complex_field, insert_before_first_anchor
│   ├── ids.py               # IdRegistry, ParaIdRegistry, _IdRegistryBase, DuplicateIdError
│   └── parts.py             # get_or_create_part, PartSpec,
│                            # COMMENTS/COMMENTS_EXTENDED/FOOTNOTES/ENDNOTES_SPEC
├── styles/                  # inspect, modify, theme
│   ├── __init__.py          # re-exports every public symbol from the submodules
│   ├── inspect.py           # resolve_effective_formatting + ResolvedFormatting + FormattingSource
│   │                        # + resolve_paragraph_spacing (contextualSpacing / the applied gap)
│   ├── modify.py            # create_style, modify_style, apply_style, delete_style,
│   │                        # ensure_style, find_matching_style, remap_styles, list_styles,
│   │                        # StyleProxy, StyleInfo, _BUILTIN_STYLES table
│   ├── sweep.py             # iter_resolved_paragraphs, ResolvedParagraph, ResolvedRun — v0.6
│   └── theme.py             # ThemeColors, load_theme, resolve_theme_color,
│                            # apply_theme_tint, apply_theme_shade, apply_lum_mod, apply_lum_off
├── controls/                # content controls (SDTs)
│   ├── __init__.py          # re-exports the public surface
│   ├── builder.py           # FormBuilder, MissingNamespaceError, DropdownItem
│   └── read.py              # ControlValue, list_controls, read_controls, set_control_value,
│                            # clear_control, WRITABLE_TYPES, ControlNotFoundError,
│                            # DuplicateTagError, ValueNotInListError, ControlTypeError
├── fields/                  # complex field insertion, reads, update flag
│   ├── __init__.py          # re-exports the public surface
│   ├── simple.py            # add_page_number_field, add_date_field, add_field,
│   │                        # PageFieldName Literal
│   ├── read.py              # read_fields, FieldInstance — v0.6
│   └── update.py            # mark_fields_dirty
├── protection/              # document-level protection enforcement
│   ├── __init__.py          # re-exports the public surface
│   └── document.py          # protect_document, unprotect_document, is_protected,
│                            # ProtectionMode Literal
├── comments/                # anchored, threaded comments — v0.2 / v0.4
│   ├── __init__.py          # re-exports the public surface
│   ├── anchor.py            # add_comment, edit_comment, delete_comment, clear_all_comments,
│   │                        # CommentRef, CommentTarget, CommentNotFoundError
│   ├── read.py              # read_comments, AnchoredComment
│   ├── threads.py           # reply_to_comment, resolve_comment, reopen_comment,
│   │                        # read_threads, CommentThread — v0.4
│   ├── _extended.py         # commentsExtended.xml thread graph (internal) — v0.4
│   ├── _ids.py              # commentsIds.xml durable ids (internal) — v0.5
│   ├── people.py            # people.xml author presence — v0.5
│   └── registry.py          # CommentIdRegistry
├── layout/                  # page-layout extras — v0.2
│   ├── __init__.py          # re-exports the public surface
│   ├── columns.py           # set_columns
│   ├── breaks.py            # insert_section_break, SectionStartType
│   ├── settings.py          # enable/disable_distinct_even_odd_headers
│   ├── line_numbering.py    # set_line_numbering, LineNumberRestart
│   └── borders.py           # set_page_borders, Border
├── bookmarks/               # bookmarks + REF/PAGEREF cross-references — v0.2
│   ├── __init__.py          # re-exports the public surface
│   ├── anchor.py            # add_bookmark, delete_bookmark, BookmarkRef, BookmarkTarget
│   ├── crossref.py          # add_cross_reference, CrossReferenceKind
│   ├── read.py              # read_bookmarks, BookmarkInfo
│   └── registry.py          # BookmarkIdRegistry
├── notes/                   # footnotes + endnotes — v0.2
│   ├── __init__.py          # re-exports the public surface
│   ├── write.py             # add_footnote, add_endnote, edit_footnote, edit_endnote,
│   │                        # FootnoteRef, EndnoteRef, NoteNotFoundError
│   ├── read.py              # read_footnotes, read_endnotes, NoteContent
│   └── registry.py          # FootnoteIdRegistry, EndnoteIdRegistry
├── numbering/               # custom list definitions — v0.5
│   ├── __init__.py          # re-exports the public surface
│   ├── define.py            # LevelDefinition, define_list_definition, define_bullet_list,
│   │                        # define_numbered_list, InvalidLevelError, MAX_LEVELS
│   ├── apply.py             # apply_list, remove_list, restart_list,
│   │                        # ListDefinitionNotFoundError
│   ├── read.py              # read_list_definitions, ListDefinition, ListLevel
│   └── registry.py          # NumIdRegistry, AbstractNumIdRegistry
├── revisions/               # tracked changes (w:ins / w:del) — v0.3
│   ├── __init__.py          # re-exports the public surface
│   ├── mark.py              # mark_insertion, mark_deletion, RevisionRef,
│   │                        # RevisionTarget, RevisionNotFoundError
│   ├── read.py              # read_revisions, TrackedChange, RevisionType
│   ├── accept.py            # accept_revision, reject_revision,
│   │                        # accept_all_revisions, reject_all_revisions
│   ├── settings.py          # enable_track_changes, disable_track_changes
│   └── registry.py          # RevisionIdRegistry
├── publishing/              # long-document publishing — v0.2
│   ├── __init__.py          # re-exports the public surface
│   ├── toc.py               # add_toc
│   ├── captions.py          # add_caption
│   ├── figures.py           # add_table_of_figures
│   └── _validate.py         # shared caption/label validation (internal)
├── tables/                  # table borders, shading, merging — v0.5
│   ├── __init__.py          # re-exports the public surface
│   ├── borders.py           # set_table_borders, set_cell_borders
│   ├── shading.py           # Shading, set_table_shading, set_row_shading,
│   │                        # set_cell_shading, shading_attrs
│   ├── merge.py             # merge_cells, unmerge_cell, normalize_horizontal_merges,
│   │                        # InvalidMergeError
│   └── read.py              # read_table_formatting, TableFormatting, CellFormatting
├── lint/                    # audit formatting, describe the repair — v0.6
│   ├── __init__.py          # re-exports the public surface
│   ├── models.py            # Finding, Issue, Location, Fix, FixOperation, Rule,
│   │                        # LintContext, RuleKind, Severity, FixOp, FixSafety
│   ├── registry.py          # rule decorator, all_rules, select_rules
│   ├── engine.py            # lint — one sweep, then every selected rule
│   ├── plan.py              # plan_fixes, FixPlan, PlannedFix, FixConflict
│   ├── profile.py           # Profile, RuleSettings, InvalidProfileError
│   └── rules/               # the twenty registered rules, by subject
│       ├── formatting.py    # style-drift, redundant-direct-formatting, ...
│       ├── typography.py    # double-space, trailing-whitespace, ...
│       ├── structure.py     # heading-level-skip, manual-list, ...
│       ├── styles.py        # unused-styles, duplicate-styles
│       ├── _common.py       # container adjacency, the one `paragraph._p` reach
│       └── references.py    # broken-cross-reference, caption-manual-numbering
├── cli/                     # docx-plus console entry point — v0.3
│   ├── __init__.py          # build_parser, main (console_scripts entry point)
│   ├── __main__.py          # python -m docx_plus.cli shim
│   ├── inspect.py           # inspect subcommand — effective formatting dump
│   ├── restyle.py           # restyle subcommand — remap_styles onto canonical ids
│   ├── controls.py          # controls subcommand — list / set / clear control values
│   ├── comments.py          # comments subcommand — list / resolve / reopen threads
│   ├── skill.py             # skill subcommand — path / list / show / install — v0.5
│   ├── lint.py              # lint subcommand — report findings — v0.6
│   ├── plan.py              # plan subcommand — describe the repair — v0.6
│   └── _io.py               # CliError + shared load/save/output helpers
├── skill/                   # packaged agent skill (Markdown, ships in the wheel) — v0.5
│   ├── SKILL.md             # entry point: frontmatter + capability map
│   └── reference/           # one topic file per capability, loaded on demand
├── examples/                # runnable demo scripts
│   ├── inspect_document.py, restyle_existing.py, build_form.py, populate_form.py
│   ├── add_comments.py, multi_column_layout.py, bookmarks_and_xrefs.py,
│   │   footnotes_and_endnotes.py     # v0.2 demos
│   ├── publishing_layout.py            # v0.2 expansion demo
│   ├── track_changes.py                # v0.3 demo
│   ├── threaded_comments.py            # v0.4 demo
│   ├── table_formatting.py, custom_numbering.py   # v0.5 demos
│   └── lint_document.py                # v0.6 demo
└── _testing/                # internal test helpers (not public API)
    ├── __init__.py
    └── ooxml_asserts.py     # assert_ids_unique, assert_style_defined,
                             # count_controls, assert_protected, assert_field_dirty
```

The flat structure is deliberate. Each capability (`styles/`, `controls/`,
…) sits as a sibling of `core/`, never deeper. There is no `_internal/`
hidden layer; `_testing/` is the only underscore-prefixed package, and it
is explicitly excluded from the public surface (`docx_plus/_testing/**`
ignores Google-docstyle in `pyproject.toml`).

The dependency rule that keeps it flat — no capability imports another —
is [invariant 1](invariants.md#the-invariants), and `cli/` and `lint/` are
its two documented exceptions.
