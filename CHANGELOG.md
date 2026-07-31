# Changelog

All notable changes to `docx_plus` are recorded here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project
uses [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- **`plan_fixes(findings) -> FixPlan`** — the linter's report-to-plan half.
  Turns findings into an ordered, inspectable, JSON-serializable
  description of what a repair pass would change, and stops there:
  **nothing in this release applies a plan.** Findings gained a `Fix`
  (`summary`, a `safety` class of `safe` / `review` / `destructive`, and a
  sequence of `FixOperation`s from a closed seven-verb vocabulary);
  `Finding.fixable` is now derived from it rather than stored, so the flag
  cannot drift from the repair.

  The planner owns the three decisions no individual rule can make,
  because each is a property of the *set* of findings:

    - **Order.** Deletions run last and back to front. Every operation
      names a position in the document as it was swept, so a deletion
      partway down invalidates every index below it.
    - **The content gate.** A fix that removes a paragraph or a style
      definition changes what the document *contains*, not how it looks.
      Those are withheld unless `allow_content=True` and reported in
      `plan.deferred` rather than silently dropped.
    - **Conflicts.** Claims are per property and per half-open character
      span, not per paragraph, so two rules clearing different properties
      of one run both apply and adjacent text edits compose. Where two
      genuinely collide the earlier wins and the loser is named.

  Every finding lands in exactly one of `fixes`, `deferred`, `conflicts`,
  or `unfixable`, so a plan accounts for the whole audit. Nine of the
  twenty rules carry a fix; the other eleven are report-only because the
  repair is a judgement the document cannot supply — promoting a heading
  or demoting the one above it are both valid, and they produce different
  documents.

- **Lint profiles** — `Profile.load(path | mapping | None)` and
  `Profile.discover(start)`, with per-rule `enabled` / `severity` /
  `options`. `lint(doc, profile=...)` applies them; both CLI commands
  discover `docx-plus-lint.json` beside the document or above it. An
  explicit `--rule` beats a profile that disabled the rule, so a
  checked-in file can never stop someone asking a direct question of one
  document. `options` is the hook policy rules will read their targets
  through and is read by nothing yet, since no policy rule ships. A
  profile naming an unknown rule raises `UnknownRuleError` on load rather
  than silently configuring nothing.

- **`docx-plus plan FILE`** — the repair, made inspectable before anything
  can apply it. Shares `--rule` / `--exclude` / `--no-tables` /
  `--profile` / `--no-profile` with `lint`, and adds `--allow-content`.
  Exits `1` when the plan holds any edit, applied or withheld, so it gates
  a pipeline on "is there anything a repair pass would do"; findings
  nobody can repair do not fail that gate. `lint` gained `--profile` /
  `--no-profile` to match.

- **`resolve_paragraph_spacing(paragraph) -> ParagraphSpacing`** — the
  vertical space Word actually leaves above and below a paragraph, as
  opposed to what the cascade declares. It folds in `<w:contextualSpacing>`
  suppression and Word's space-after/space-before arithmetic (below), so
  one paragraph's `space_below` always equals the next one's `space_above`.
  Alongside those two it reports `declared_before` / `declared_after`, the
  resolved `contextual_spacing` flag, and which edges were suppressed.

### Fixed

- **`double-space` missed every other occurrence in a paragraph.** The
  pattern was `\S {2,}\S`, which consumes the word between two runs of
  spaces, so the search for the next one started past the character that
  would have anchored it: `"a  b  c"` reported one occurrence, not two.
  Invisible while the rule only reported — one finding per paragraph
  either way — and it would have left the paragraph half-repaired the
  moment a fix carried the spans. Now anchored with lookarounds.

- **The default paragraph style was never applied, so a paragraph with no
  `w:pStyle` resolved to `docDefaults` alone.** That is most paragraphs in
  most documents: everything `Normal` declares — its font, its size, its
  spacing — was silently dropped, and `style_id` came back `None` where
  Word reports `Normal`. Measured against live Word:

    - The default style is the **last** `w:style w:type="paragraph"` whose
      `w:default` is on (`1` / `true` / `on`), falling back to the style
      whose id is literally `Normal`. Declaration order is the tie-break,
      measured both ways round.
    - It substitutes whenever `w:pStyle` fails to resolve — **absent,
      dangling, or naming a style of the wrong `w:type`** — and then acts
      as an entirely ordinary paragraph style: it is the reported
      `style_id` / `style_name`, it supplies numbering, and it counts as
      one of the toggle rule's *levels*, so a bold `Normal` and a bold
      character style cancel.
    - It sits at the **paragraph-style layer**, which is what the
      measurement pinned down: a `Normal` declaring 20pt beats a table
      style declaring 36pt, and beats the table style's `w:tblStylePr`
      branches too. Collapsing it into `docDefaults` would have matched
      every other reading and got this one wrong.
    - The other two `w:default="1"` styles are **non-events**:
      `DefaultParagraphFont` never reaches a run and `TableNormal` never
      reaches a table naming no style. Only `w:pStyle` has a fallback.

  A `w:pStyle` naming an undefined style now reports the default style
  rather than the name it wrote — Word renders it that way, and code
  keying off `style_id` was matching a style the document does not have.

- **Style references were followed without checking `w:type`.** Word
  resolves a reference only to a style of the type it demands, and ignores
  it otherwise. The resolver followed every one: a `w:rStyle` naming a
  paragraph style, a `w:tblStyle` naming a paragraph style, and a
  `w:basedOn` crossing between paragraph and character styles all
  contributed formatting Word never applies. A cross-type `w:basedOn` now
  ends the chain — the style itself still applies, it just inherits
  nothing through the dead link.

- **A table cell resolved as `docDefaults` plus the table style and
  nothing else.** The cell cascade now ends with the default paragraph
  style, which is what a bare paragraph in the cell picks up and what Word
  reports for an untouched cell — so a cell in a styled table reports the
  default style's size beating the table style's, rather than the table
  style's alone.

- **`<w:contextualSpacing>` was ignored completely, so every list
  paragraph in a stock-template document got the wrong spacing.** Fourteen
  of Word's built-in styles carry the flag — `ListParagraph`, `Title` and
  every `List*` / `ListBullet*` / `ListNumber*` — and the resolver reported
  their declared space between adjacent items where Word renders none.
  Measuring it turned up a second divergence nobody was looking for:

    - **Word does not add space-after to the next paragraph's
      space-before.** It lays down the space-after and then tops it up to
      the space-before if that is larger, so an ordinary pair sits
      `max(after, before)` apart, not `after + before`. Anything computing
      a gap by summing the two resolved values was overstating it, with or
      without `contextualSpacing` involved.
    - **The suppression removes one of those two terms, not the whole
      gap.** Each edge answers only to its own paragraph's flag, and the
      top-up is still measured from the **declared** space-after even when
      that space-after was itself suppressed — a contextual paragraph with
      20pt after followed by a non-contextual one with 30pt before leaves
      10pt, not 30pt.
    - **"Same style" means `styleId` identity and nothing else.**
      Numbering plays no part: two `ListParagraph` paragraphs in unrelated
      lists, or at different levels of one list, suppress exactly as two
      in the same list do. A `basedOn` child is a *different* style, even
      though it inherits the flag.
    - **A content control is transparent.** A `<w:sdt>` wrapping the
      neighbour, or holding a paragraph between the pair, leaves the
      paragraphs adjacent. A table between them does not.

  `ResolvedFormatting` gains `contextual_spacing`, resolved through the
  ordinary cascade (`docDefaults` included) as a plain override rather than
  an ECMA-376 17.7.3 toggle. `spacing_before` / `spacing_after` are
  deliberately **unchanged** and still report what the cascade declares —
  whether either is applied depends on the paragraph's neighbours, and the
  linter's `style-drift` rule compares direct formatting against a style,
  which needs the declared numbers. The new
  `resolve_paragraph_spacing` answers the layout question.

  Measured across 111 gaps read out of Word's own layout, including a full
  4×4×2×2 sweep of space-after × space-before × the flag on each
  paragraph; the resolver now matches every one. As with the theme work,
  COM could not supply the ground truth — `ParagraphFormat.SpaceBefore`
  reports what the cascade says, which is the thing in dispute — so the
  probes were exported to PDF and the paragraph baselines measured.

- **Theme colours ignored the document's `<w:clrSchemeMapping>`, and the
  tint/shade arithmetic did not match Word.** Two independent problems,
  both found by exporting a probe document from Word and reading the
  colours it actually wrote:
    - **`<w:clrSchemeMapping>` is now honoured.** `settings.xml` decides
      which theme slot a `w:themeColor` name resolves to; the resolver
      hardcoded `text1`→`dk1`, `background1`→`lt1` and so on. That matches
      Word's default mapping, so ordinary documents were fine — but a
      document that remaps the slots (a dark-themed template swaps `t1`
      and `bg1`) resolved `text1` to **black where Word renders it
      white**, and every accent and hyperlink colour could be redirected
      too. Only the semantic names follow the mapping; `dark1` / `light1`
      / `dark2` / `light2` name a slot outright and are never redirected,
      so this is not a rename of the scheme. `ThemeColors` gains a
      `mapping` field, defaulting to Word's defaults so an absent or
      partial element behaves as Word treats it.
    - **`themeTint` / `themeShade` now use exact arithmetic.** The
      transforms ran through `colorsys` in binary floating point, which
      changes the answer at the integer boundaries these values keep
      landing on: `1 - 0xE6/255` is 0.09803921568627449, and 255 times
      that is 24.999999999999996 — one below the 25 Word paints. Against
      47 measured tint/shade values the resolver was exact on 18 with a
      worst-case error of 2 units per channel; it is now exact on 32 with
      a worst case of 1. Exactness also makes the RGB→HSL→RGB round-trip
      lossless, which is what lets the final channel be truncated (as
      Word does) while `themeTint="FF"` stays a no-op.

  The residual is real and enumerated rather than waved at:
  `tests/test_theme_word_verified.py` lists the 15 cases that land within
  one unit of Word without matching exactly, and fails if a case is added
  to that list or if a listed case starts matching. Word's rounding at
  those boundaries was not reverse-engineered.

  Two knock-on changes: `apply_theme_shade("4F81BD", "80")` now returns
  `244061`, which is what Word renders — the previous `254062` was
  asserted as Word-verified but nothing renders it. And `apply_lum_mod` /
  `apply_lum_off` shift by up to one unit as a consequence of sharing the
  arithmetic. Those two remain **unverified** against Word, and cannot be
  verified through this library: they are DrawingML transforms and
  `w:color` has no attribute that carries them, so no cascade input
  produces one. The documentation previously described them as merely
  "not wired into the cascade walker yet", which implied a gap where
  there is none.

- **Conditional table-style formatting ignored `<w:tblLook>` entirely,
  and is now measured rather than inferred.** The resolver picked
  `<w:tblStylePr>` branches from the cell's position alone, so it applied
  header-row, first-column and banding formatting to tables that had those
  boxes unticked. Across 1015 cells read back from live Word it agreed on
  19 of the first 80; it now agrees on all 1015. Six distinct causes, each
  a behaviour change:
    - **`w:tblLook` now gates every branch.** These are the Header Row /
      First Column / Banded Rows tick-boxes in Word's Table Design tab.
      Both forms are honoured — the named attributes and the legacy
      `w:val` hex bitmask that Word 2007 wrote — and a table with *no*
      `<w:tblLook>` enables everything, which is what Word does rather
      than the `0`-defaults the schema implies. Because python-docx's
      `add_table` emits Word's default look, `lastRow`, `lastColumn` and
      vertical banding are off unless a document asks for them.
    - **A corner branch needs both of its axes.** With `firstColumn`
      cleared, the top-left cell takes `firstRow`, not `nwCell`.
    - **Banding requires a declared band size.** `w:tblStyleRowBandSize` /
      `w:tblStyleColBandSize` is read from the table instance and then its
      style chain — previously only the instance was consulted, and an
      absent value defaulted to `1`. Absent means **no banding**; Word's
      own table styles all declare an explicit `1` for this reason.
    - **The band sequence starts at row / column 0.** It shifts to 1 only
      when the matching `firstRow` / `firstCol` conditional actually
      paints that line — the `tblLook` flag alone does not shift it. The
      old code always skipped line 0, which was right only for the common
      header-row case.
    - **Branch precedence was inverted on two axes.** A vertical band
      beats a horizontal one, and a row branch beats a column branch. Both
      are the opposite of the order ECMA-376 17.7.6.5 lists, which is what
      the old code followed — and document order is no guide either, since
      Word rewrites a style's branches into the spec's order on save.
    - **The `wholeTable` branch is inert.** Word discards
      `<w:tblStylePr w:type="wholeTable">` on load: it renders neither the
      `rPr` nor the `pPr` and drops the element when it saves. Whole-table
      formatting belongs on the style's own `w:rPr` / `w:pPr`, which the
      base pass already applies.

  `TableContext` gains `first_row_enabled`, `last_row_enabled`,
  `first_col_enabled` and `last_col_enabled`, all defaulting to True so a
  hand-built context behaves like a table with no `<w:tblLook>`. The
  positional fields keep their meaning.

  `tests/test_tables_word_verified.py` carries the measured grids as a
  parametrised table so the behaviour cannot drift back.

- **The cascade resolver's toggle rule was wrong, and is now measured
  rather than inferred.** Bold, italic and the other ten ECMA-376 17.7.3
  toggles were folded with a running XOR over the whole walk. Checked
  against live Word across 79 probe cases, the resolver agreed on only 51
  of 74 property answers. It now agrees on all of them. Six distinct
  causes, each a behaviour change:
    - **`basedOn` chains no longer alternate.** A chain is *one* level of
      the hierarchy and flattens by plain override, so a child style
      restating its parent's `<w:b/>` stays bold. Previously every such
      pair cancelled — the single most common shape in a real
      `styles.xml`.
    - **Direct formatting is absolute.** A bare `<w:b/>` on a run states
      bold rather than flipping it, so direct bold over a bold style
      resolves bold. Word writes `w:val="0"` when a user un-bolds
      something, which is the case that genuinely turns it off.
    - **`docDefaults` is the base, not a level.** A style restating the
      document default is inert instead of cancelling it.
    - **`w:val="0"` at a style level is only meaningful when it differs
      from the base.** A table style asking for bold and a paragraph
      style asking for not-bold resolves *bold*, which is what Word
      renders.
    - **The `linkedCharStyle` layer is gone.** A paragraph style's
      `w:link` partner is a UI affordance for applying its character half
      to a selection; Word never consults it to render a run. A style
      whose formatting lived only on its `Char` half used to resolve, and
      no longer does. `Layer` loses the member — **breaking** for callers
      naming it in a `stop_below` argument or matching on
      `FormattingSource.layer`.
    - **A numbering level's `w:rPr` no longer reaches run text.** It
      formats the number or bullet glyph, so a level carrying `<w:b/>`
      renders a bold "1." in front of unbolded prose. Reporting it as the
      paragraph's own formatting described something the reader never
      sees.

  This moves two `lint` rules in the user's favour: `redundant-direct-
  formatting` now correctly flags direct formatting that restates an
  already-resolved toggle, which it previously protected on the false
  premise that removing it would change the rendering.

  `tests/test_cascade_word_verified.py` carries the measurements as a
  parametrised table so the rule cannot drift back.

### Added

- **Community health files** — `CONTRIBUTING.md`, `SECURITY.md`,
  `CODE_OF_CONDUCT.md` (Contributor Covenant 2.1), YAML issue forms for
  bug reports and feature requests, and a pull request template. The
  contributing guide records the release process and the
  verify-against-Word expectation that had until now been folklore; the
  security policy is explicit that `protect_document` is an editing
  convention and **not** a security boundary, and that the realistic
  XML/ZIP attack surface belongs to `lxml` and `python-docx` upstream.
- **`Changelog` project URL** — PyPI renders it as a sidebar link.
  `CHANGELOG.md` and `CONTRIBUTING.md` now ship in the sdist.

### Changed

- **README rewritten for new users.** It opened with a fifteen-bullet
  feature list and its only install instructions were "Install
  (development)" — clone the repo and `uv sync`. Six versions have been
  on PyPI since v0.1.0 and `pip install docx-plus` appeared nowhere,
  which is the one thing a reader arriving from the package page needs.
  It now leads with the problem the library solves, a runnable example,
  and the install command; the feature list is a capability table
  linked into the architecture docs; and the quickstart keeps the five
  most-used surfaces rather than reprinting all twelve. Repo-relative
  links are absolute GitHub URLs so they resolve on PyPI too.
- **Stale version claims corrected.** The capability list was headed
  "v0.1 through v0.4" and `docs/index.md` stopped its roadmap table at
  v0.4, so v0.5's `tables/` and `numbering/` were invisible outside the
  collapsed build-history block. The "What's next" section still
  described v0.2 in the present tense.
- **Package metadata** — keywords expanded from six to fifteen to cover
  the differentiating surfaces (content controls, tracked changes,
  footnotes, bookmarks), and five classifiers added
  (`Environment :: Console`, `Topic :: Text Processing :: Markup :: XML`,
  and others).
- **Docs site** — `repo_url` / `edit_uri` set, so Material renders the
  GitHub link and per-page "edit" actions; changelog and contributing
  guide added to the nav.

- **`lint/` — the document linter.** A new *composing* layer: like `cli/`
  it sits above the capability modules and reads across them, adding no
  OOXML knowledge of its own. `lint(doc)` audits a document and returns
  `Finding`s; nothing is modified, and applying fixes is deliberately not
  in this release.
  Rules divide into three **kinds**, which is what keeps an opinionated
  feature inside a lean library. A `consistency` rule reports that a value
  fights the document's *own* applied styles — the document supplies the
  target, so no opinion is imposed. A `structural` rule reports an
  objective defect. A `policy` rule compares against a target *you*
  supply, and none ship enabled, which a test enforces.
  Nine rules to start: `double-space`, `trailing-whitespace`,
  `space-before-punctuation`, `indent-by-whitespace`,
  `stray-empty-paragraph`, `heading-level-skip`, `empty-heading`,
  `manual-list`, `redundant-direct-formatting`, and `mixed-run-formatting`.
  The last two are where resolving OOXML beats asking Word: `w:rStyle`,
  numbering-level `rPr`, and table-style conditional formatting are all
  invisible behind an effective value, and `FormattingSource` names the
  exact layer. `manual-list` is only possible at all because numbering now
  resolves through the style chain.
  Rules register by decorator, so adding one is a single function.
  Selection is by rule id **or tag**, where naming a tag also enables that
  cluster's off-by-default rules; an unknown selector raises rather than
  silently matching nothing, since "no rules ran" and "no findings" look
  identical otherwise.
- **`resolve_effective_formatting(..., stop_below=Layer)`** — resolve with
  a cascade layer, and everything above it, excluded. This answers a
  question provenance cannot: provenance names the layer that *won*, not
  the value that would have surfaced in its absence. Resolving a run with
  `stop_below="directRun"` gives exactly what it would render as if its
  own `<w:rPr>` were deleted — character style and all.
  `style_id` / `style_name` are identity rather than formatting, so they
  are reported wherever the walk stops. The two numbering layers gate
  separately, which is what the `styleNumbering` / `numbering` split was
  for: `stop_below="numbering"` reports the list a paragraph's *style*
  would give it, ignoring its own `w:numPr`.
- **`iter_resolved_paragraphs(..., include_baseline=True)`** — populates
  `.baseline` on each `ResolvedParagraph` and `ResolvedRun` with that
  target resolved one layer down (`directParagraph` / `directRun`
  respectively), sharing the sweep's cache. Off by default, since it
  roughly doubles the resolve work.
- **`style-drift`** — a lint rule for paragraph-level direct formatting
  that deviates from the applied style. The counterpart to
  `redundant-direct-formatting`: the same comparison against the same
  baseline, split because the two imply opposite actions. It is the
  sibling COM linter's central rule and the clearest case for resolving
  OOXML instead: a two-layer compare sees a numbered paragraph's
  level-supplied indent as drift, because it cannot tell which layer
  produced the effective value.
- **`LintContext.resolve()`** — an escape hatch for a rule needing a slice
  of the cascade the sweep did not precompute. Not cache-shared, and
  documented as such.
- **`fields.read_fields`** — the read half of `fields/`, which until now
  could only write. A complex field is a *run sequence* delimited by
  `w:fldChar` markers, with its instruction split across however many
  `w:instrText` elements Word chose, so reading one back is a walk rather
  than an xpath. `FieldInfo` splits the instruction into `keyword`,
  `operands`, and `switches`; nested fields (which Word writes for `TOC`
  and `IF`) read as one field under the outer keyword.
- **`ResolvedFormatting.lang`** — the `w:lang` Latin-script language tag,
  resolved through the cascade like everything else. Only `w:val` is
  surfaced; `w:eastAsia` and `w:bidi` are separate properties for separate
  scripts, and collapsing three languages into one field would misreport
  which one a proofing tool uses.
- **`broken-cross-reference`, `caption-manual-numbering`, and
  `mixed-language`** — the last three rules of the v0.6 table.
  `broken-cross-reference` is the catalogue's only `error`: a `REF` to a
  missing bookmark renders as *Error! Reference source not found.* the
  moment fields recalculate, and until then shows the stale cached result,
  which is how it goes unnoticed. Both field rules read the instruction
  rather than that result. `mixed-language` compares against the
  document's own majority tag, so there is nothing to configure.
- **`styles.find_unused_styles`** — the read companion to `delete_style`:
  which definitions could be removed without breaking a reference. Usage
  is a **closure**, not a single pass, since a style referenced only by
  another unused style is itself unused. An unused `w:link` pair collapses
  to its paragraph half, because `Heading 1 Char` exists only because
  `Heading 1` does.
- **`StyleInfo.is_builtin`** — whether a style came from the template or
  was authored in the document, read from `w:customStyle` (ECMA-376
  17.7.4.9). Chosen over the known-built-ins table because that table does
  not cover the table-style gallery, which is most of what a stock
  template ships.
- **`duplicate-styles` and `unused-styles`** — the first rules whose
  subject is a style definition rather than a paragraph, so their findings
  carry a `style_id` and (for `unused-styles`) no position at all.
  `duplicate-styles` compares each paragraph's **baseline**, so two styles
  reaching the same formatting by different `basedOn` routes still match
  and the author's direct overrides do not confuse it.
- **Four more lint rules.** `direct-numbering-override` (a paragraph's own
  `w:numPr` fighting the list its style supplies — the rule `stop_below`
  was needed for, and it downgrades `numId=0` to `info` since the opt-out
  sentinel is the one legitimate override); `list-numbering-continuity`
  (the "three separate 1. lists" footgun — adjacent items at one level
  belonging to different `numId`s); `manual-heading-formatting` (a bold or
  enlarged short line standing in for a heading, so it is missing from the
  navigation pane and every generated TOC); and `font-outliers` (off by
  default — thinly-populated font/size combinations against the
  document's dominant set). `manual-heading-formatting` and
  `font-outliers` compare against the *document's own* body size and
  dominant font rather than a fixed threshold, which is what keeps them
  consistency judgements rather than house opinions.
- **`docx-plus lint`** — the CLI over it. `--rule` / `--exclude` take an
  id or a tag and repeat, `--list-rules` prints the catalogue without
  needing a document, `--json` emits the full finding shape, and
  `--no-tables` skips table cells. It exits **`1` when it finds
  something**, so it works as a CI gate; a genuine failure is still
  distinguishable by its `error:` line on stderr.
- **`styles.iter_resolved_paragraphs`** — a document-wide cascade sweep,
  resolving every paragraph and run against one shared cache and yielding
  lazily in document order. `resolve_effective_formatting` answers "what
  does *this* paragraph render as", rebuilding the theme, the styles part,
  and each `basedOn` chain on every call; asking it about a whole document
  re-does that work per target, and profiling put `load_theme` alone at
  39% of per-call cost. Sharing those lookups is **~5x faster per
  target** (0.61 ms → 0.13 ms on a 1500-paragraph document), which is what
  makes whole-document analysis practical.
  Results come back as `ResolvedParagraph` (paragraph, index, formatting,
  runs, `table_depth`) and `ResolvedRun`. The walk uses python-docx's
  `iter_inner_content`, so paragraphs and tables stay interleaved in
  document order — `doc.paragraphs` drops the tables and `doc.tables`
  drops the ordering. Nested tables are descended, and a merged cell is
  visited **once** rather than once per grid column it spans, which a
  naive `row.cells` walk gets wrong. `include_runs=False`,
  `include_tables=False`, and `include_provenance=True` tune the cost.
  Headers, footers, and notes are out of scope for now; only the main
  document body is swept.

### Fixed

- **A run inside a bold `Heading 1` resolved `bold=False`.** Word writes a
  paragraph style and its `w:link` character partner with identical
  `w:rPr`, and the resolver applied both as independent levels of the
  style hierarchy. Toggle properties XOR between levels per ECMA-376
  17.7.3, so every toggle the pair agreed on cancelled itself out: a stock
  `Heading 1` **paragraph** resolved `bold=True` while a **run** inside it
  resolved `bold=False`. A `w:link` partner is not a further level — it is
  the same style's character half — so its toggles now state a value
  rather than flipping one. The layer still overrides normally, which is
  what lets a style carrying its character formatting solely on the Char
  half resolve at all.
  Found while building the lint baselines, which made the paragraph and
  run answers directly comparable for the first time; a test now asserts
  the two targets agree about what a style says.
- **A run's baseline no longer inherits paragraph-mark formatting.** The
  `rPr` inside a `pPr` formats the pilcrow, not the runs, and is correctly
  excluded from a run-target resolve. `redundant-direct-formatting`
  previously compared runs against the *paragraph's* resolve, so a
  paragraph mark carrying a size made a run matching it look redundant
  when deleting the direct property would have changed the rendering.
- **`resolve_effective_formatting` now resolves style-supplied
  numbering.** Cascade layer 4 read only the paragraph's *direct*
  `w:numPr`, so on a stock `Document()` a paragraph styled
  `List Bullet` reported `num_id=None` even though the bundled template
  links `numId` 1 on that style. Every other field on
  `ResolvedFormatting` walks the style chain, so `num_id` silently broke
  the contract the rest of the dataclass sets — and it made a
  correctly-styled list paragraph indistinguishable from one where a
  bullet glyph was typed by hand. Numbering now resolves through the
  `basedOn` chain, nearest style winning. Reaching the definition at all
  also means the `abstractNum` level's own `pPr` applies, so such a
  paragraph picks up its indent for the first time.
  `w:numId` and `w:ilvl` resolve **independently**: a paragraph
  overriding only the level keeps its style's list rather than losing
  numbering. ECMA-376 does not state merge semantics for a compound
  property across the style / direct boundary, so this was settled
  against Word 2016 — a `List Bullet` paragraph given a bare
  `<w:ilvl w:val="2"/>` renders as a third-level bullet of the style's
  own list.
  A resolved `num_id` of **`0` is surfaced rather than flattened to
  `None`**: it is the ECMA-376 17.9.18 "explicitly not numbered"
  sentinel, the only way to opt out of numbering a style applies and
  what `numbering.remove_list(suppress_style_numbering=True)` already
  writes, so `None` now means "no numbering information anywhere" and
  `0` means "deliberately suppressed".
- **New `styleNumbering` provenance layer** on `FormattingSource`,
  carrying the `style_id` that supplied the reference and its
  `chain_depth`. A distinct layer rather than a `numbering` entry with a
  `style_id` set, because `Layer` describes where in the cascade a value
  sat and a style's `numPr` is overridden by a direct one. The
  formatting the numbering *level* contributes stays `numbering`, whose
  precedence is the same either way.

## [0.5.0] - 2026-07-27

### Added

- **The agent skill now ships in the wheel.** It lived at repo-level
  `skills/docx-plus/` through v0.4, which meant `docs/SKILLS.md` claimed
  the library "ships" it while linking only to GitHub blob URLs —
  broken for anyone who had `pip install`ed rather than cloned. The tree
  moved to `docx_plus/skill/`, which needed **no build configuration
  change at all**: hatchling's `packages = ["docx_plus"]` already sweeps
  non-`.py` files, the same reason `py.typed` ships. Verified by
  building a wheel, unzipping it, and driving the result from a clean
  venv with no source tree.
- **`docx-plus skill`** — a new CLI command over that:
  `install [--dest DIR | --user] [--force]` copies the tree into a
  skills directory (defaulting to `./.claude/skills/`), and
  `path` / `list` / `show [TOPIC]` locate or read it in place. The one
  command that touches no `.docx`, so the `-o/--output` / `--in-place`
  convention does not apply to it.
- **Comment durable ids (`commentsIds.xml`)** — a comment has three
  identifiers and only this one survives an edit: `w:id` is a
  position-dependent index Word renumbers, and the `w14:paraId` the
  thread graph keys off changes whenever the body is rewritten. Word
  2016 added the part for exactly this reason, and anything citing a
  comment from outside the document — a permalink, a review tracker, a
  diff between two revisions — needs it.
  `add_comment` and `reply_to_comment` now write an entry (Word
  regenerates missing ones anyway, so emitting them moves output toward
  native); `delete_comment` and `clear_all_comments` remove it; the
  value surfaces as `AnchoredComment.durable_id`. `DurableIdRegistry`
  allocates. Both take an optional `durable_id_registry` to share across
  a batch.
  **The id is hex, not decimal** — `ST_LongHexNumber`, the same
  8-uppercase-digit form as `w14:paraId`. That was established by having
  Word 2016 author a commented file and reading it back, which also
  confirmed all four content-type and relationship URIs.
- **Comment author presence (`people.xml`)** — `set_author_presence` /
  `read_author_presence` / `clear_author_presence` over a new
  `AuthorPresence` record, writing the `<w15:person>` /
  `<w15:presenceInfo>` pair that drives the presence indicator in Word's
  reviewing pane.
  **Opt-in by design:** `add_comment` does *not* write this part.
  Registering an author means inventing a `userId` for someone the
  library knows nothing about, and a fabricated directory identity is
  worse than an absent one. The part is cosmetic — comments, threading,
  and resolution all work without it. Stale authors are not pruned on
  delete, matching Word.
- **`_testing.assert_durable_ids_well_formed`** — checks that every
  `durableId` and `paraId` in `commentsIds.xml` is unique and is 8
  uppercase hex digits.
- **Table formatting (`tables/`)** — the half of tables python-docx
  omits. It models structure well — rows, columns, cells, widths, and a
  working `_Cell.merge` — and appearance not at all: there is no
  `CT_Border`, `CT_TblBorders`, `CT_TcBorders`, or `CT_Shd` class in
  the package and none of those tags is registered, so ruling a table
  or shading a header row has meant writing OOXML by hand.
  `set_table_borders` / `set_cell_borders` write `<w:tblBorders>` and
  `<w:tcBorders>` over the same `Border` dataclass page borders use,
  including the inside edges and the cell diagonals; both pin
  `w:space` to `0`, because `Border`'s default of `24` is a *page*
  value Word's table UI cannot even produce. `set_table_shading` /
  `set_cell_shading` / `set_row_shading` write `<w:shd>` — the last of
  those writes through to every cell, because `CT_TrPr` has no
  shading child and Word does the same.
  `merge_cells` wraps python-docx's own merge in a typed error rather
  than re-implementing it; `unmerge_cell` is the inverse, which
  python-docx lacks entirely — nothing in it removes a `w:gridSpan` or
  a `w:vMerge`, so a merge was one-way. `normalize_horizontal_merges`
  rewrites legacy `<w:hMerge>` spans as `<w:gridSpan>`, the only form
  python-docx's grid model understands; verified rendering-preserving
  against Word 2016, where the converted file rasterises byte-for-byte
  identically. `read_table_formatting` reads it all back — direct
  formatting only, since the table-style cascade stays out of scope.
  A new `docx_plus/examples/table_formatting.py` builds a ruled,
  shaded, merged budget table.
- **Custom numbering (`numbering/`)** — the largest remaining
  python-docx gap. python-docx ships a `NumberingPart` and `len()` of
  its definitions and nothing else: it has no `CT_AbstractNum` and no
  `CT_Lvl`, so it cannot express a number format, level text, start
  value, indent, or bullet glyph. Building a list has meant hand-writing
  XML.
  `define_list_definition` writes a `<w:abstractNum>` from a list of
  `LevelDefinition`s plus the `<w:num>` instance that paragraphs
  reference; `define_bullet_list` and `define_numbered_list` are presets
  using Word's own glyph and format cycles. `apply_list` /
  `remove_list` set paragraph membership — the latter's
  `suppress_style_numbering` writes the `numId="0"` sentinel, the only
  way to opt out of a list a *style* applies. `restart_list` begins a
  fresh sequence, which OOXML has no paragraph-level way to express: it
  adds a second `<w:num>` over the same `<w:abstractNum>` with a
  `<w:startOverride>`, exactly as Word does. `read_list_definitions`
  reads the part back as `ListDefinition` / `ListLevel`, including
  definitions other tools wrote.
  A new `docx_plus/examples/custom_numbering.py` builds a procedure, a
  restarted sequence, a legal outline, and a three-level bullet list;
  verified against Word 2016.
- **`fields.add_style_reference`** — the `STYLEREF` field, which resolves
  to the text of the nearest paragraph carrying a given style and is the
  one cross-reference that needs no bookmark at all. Word re-evaluates it
  per page, so a single field in a header gives a running chapter title.
  Takes the style *name* as Word shows it (`"Heading 1"`) rather than the
  `w:styleId`, because that is what the field instruction accepts, or an
  `int` outline level. Reachable before this only as a raw
  `add_field(instruction=...)` string with no validation.
- **Caption cross-references** — `publishing.add_caption` gained
  `bookmark_name`, which brackets the label and `SEQ` number in a
  bookmark. This is what makes "see Figure 3" possible: a `REF` field
  **cannot point at a `SEQ` field**, only at a bookmark, and captions
  previously created none — so the reference had nothing to target. The
  bookmark spans exactly the extent Word's own "Only label and number"
  option uses, so a bare `add_cross_reference` to it resolves to
  `"Figure 3"`; description text added afterwards stays outside.
- **`bookmarks.add_cross_reference` gained the rest of the switch
  surface** — `number` (`\n` / `\r` / `\w`, the target's paragraph
  number with varying context), `position` (`\p`, resolving to
  `"above"` / `"below"`), `suppress_non_delimiters` (`\t`),
  `numeric_format` (`\#`), and `preserve_formatting`
  (`\* MERGEFORMAT`). It also now validates `bookmark` against Word's
  name grammar, since a name only Word's UI would reject produced a
  silently unresolved field, and rejects the `REF`-only switches when
  paired with `kind="page"`.
- **`core.BookmarkNameRegistry`** — bookmarks are the one thing in the
  format addressed by name, and nothing stopped a document carrying two
  with the same `w:name`, which makes a `REF` ambiguous and makes
  `delete_bookmark` remove both. `next_ref_name()` mints hidden anchors
  in Word's own `_Ref` + 9-digit form; the leading underscore is what
  keeps them out of Word's Bookmark dialog.
- **`core.validate_bookmark_name`** — the name grammar, now shared by the
  three surfaces that accept a bookmark name instead of living privately
  in `bookmarks/anchor.py`.
- **`core.ordered_insert`** — idempotent schema-ordered insertion given a
  parent's full child sequence. Promoted out of `styles/modify.py` so the
  numbering writer can share it; SPEC §9.1 forbids that sibling import.
- **`core.build_bookmark`** — the `bookmarkStart` / `bookmarkEnd` emitter,
  extracted from `bookmarks/anchor.py` for the same reason: `publishing`
  needs to bookmark a caption, and a `REF` field can only point at a
  bookmark, never at the caption's own `SEQ` field.
- **`_IdRegistryBase.next_sequential()`** — lowest-free-integer
  allocation, alongside the existing random `next()`. Word and python-docx
  both number lists this way, and `reserve()` now honours a per-registry
  `_MIN_ID` so `w:abstractNumId`, which legitimately starts at 0, can use
  the shared machinery.
- **`core.NUMBERING_SPEC`, `COMMENTS_IDS_SPEC`, `PEOPLE_SPEC`** — part
  specs for `/word/numbering.xml`, `/word/commentsIds.xml`, and
  `/word/people.xml`, with the `w16cid` namespace and the `CT_` / `RT_`
  URI constants for the two Microsoft extension parts.
- **`_testing.assert_numbering_well_formed`** — checks the three
  invariants a lenient parser will not: that every `w:abstractNum`
  precedes every `w:num` (`CT_Numbering`'s child order, which nothing in
  python-docx maintains because nothing in it inserts an `abstractNum`),
  that both id namespaces are unique, and that every instance resolves.

### Changed

- **`_IdRegistryBase.next_hex()`** — the `ST_LongHexNumber` rendering
  moved up from `ParaIdRegistry`, since `w16cid:durableId` uses the
  identical form. `ParaIdRegistry.next_hex` is unchanged for callers.
- **`BookmarkIdRegistry` moved from `docx_plus.bookmarks.registry` to
  `docx_plus.core.ids`** and is re-exported from its old home, so
  existing imports are unaffected. Forced by SPEC §9.1: `publishing` needs
  it to bookmark a caption and cannot import from a sibling capability.
- **`Border` moved from `docx_plus.layout` to `docx_plus.core.borders`**,
  where table and cell borders can share the identical `CT_Border`
  shape. `docx_plus.layout.Border` re-exports it, so existing imports
  are unaffected. It also gained the `size` range check and `style`
  validation its docstring had claimed since v0.2 but never enforced.
- **`MissingPartError` is documented as unraised.**
  `resolve_effective_formatting` promised it for an unresolvable
  numbering reference; nothing ever raised it, and the tolerant
  behaviour is correct — Word degrades the same way. The docstring now
  says so. The exception is retained as a public symbol.

### Fixed

- **`resolve_effective_formatting` no longer crashes on a document with
  no `numbering.xml`.** `doc.part.numbering_part` fabricates a missing
  part through `NumberingPart.new()`, which is an unimplemented stub in
  python-docx that raises a bare `NotImplementedError`. The cascade's
  `_numbering_root` reached it through `getattr(..., None)`, which
  swallows only `AttributeError`, so any paragraph carrying a `w:numPr`
  without the part behind it raised — the state LibreOffice, Pandoc, and
  stripped templates all produce. The resolver now reads the
  relationship directly and treats an absent part as "no numbering
  information", which is what it already did for a dangling `numId`.

## [0.4.0] - 2026-07-26

### Added

- **Threaded comments (`comments/threads.py`)** — the reply / resolve model
  Word has used since 2013, backed by the `commentsExtended.xml` part
  neither python-docx nor `docx_plus` previously wrote.
  `reply_to_comment` attaches a reply beneath an existing comment,
  mirroring the parent's anchor range so Word renders the thread as one
  balloon, and appending its range markers after every marker the thread
  already owns — Word orders a thread's balloons by marker position in
  the body, so that placement is what makes replies read in conversation
  order; `resolve_comment` / `reopen_comment` toggle a thread's
  resolved state (`w15:done`), thread-wide as in Word's UI, so naming any
  member moves the whole thread; `read_threads` returns comments grouped
  as `CommentThread` (root, replies, resolved). `read_comments` results
  gained `parent_id` and `resolved`. `delete_comment` gained
  `include_replies` (default `True`), which deletes a root's reply
  subtree the way Word does. A new
  `docx_plus/examples/threaded_comments.py` and the refreshed `comments`
  agent-skill reference round out the surface.
- **`docx-plus comments`** — `list` (with `--unresolved` and `--json`,
  printing replies indented under their root), plus `resolve` and
  `reopen`, which require `-o/--output` or `--in-place` per the CLI's
  mutating-command convention.
- **`core.ParaIdRegistry`** — allocator for `w14:paraId`, the id threaded
  comments key their parent/child links off. Unlike every other registry
  it is unique across the whole *package*, so it seeds from the document
  body plus the comments / footnotes / endnotes parts; `next_hex()`
  renders the 8-uppercase-hex-digit form Word writes.
- **`core.COMMENTS_EXTENDED_SPEC`** — `PartSpec` for
  `/word/commentsExtended.xml`, alongside the `CT_COMMENTS_EXTENDED` /
  `RT_COMMENTS_EXTENDED` URIs (Microsoft extensions with no member in
  python-docx's `CT` / `RT` enums) and an `XmlPart` registration so an
  existing extended part round-trips as parsed XML rather than a blob.
- **`core.ns.W15` and `core.ns.BUILD_NSMAP`** — the Word 2012 namespace,
  and a narrower construction map that keeps `el()` from declaring the
  extension prefix on every element written into `document.xml`.

### Changed

- **`add_comment` now writes thread metadata.** Every comment it inserts
  is stamped with a `w14:paraId` and registered as an unresolved root in
  `commentsExtended.xml`, which is created on first use — matching what
  Word writes, and making any comment reply-able and resolvable the
  moment it exists. Documents that previously carried only
  `comments.xml` will gain the second part. `add_comment` accepts an
  optional `para_id_registry` for sharing an allocator across a batch.
- **`edit_comment` preserves the thread link.** The rebuilt body
  paragraph carries the comment's original `paraId`, so replies stay
  attached and a resolved thread does not silently reopen.
- **`clear_all_comments` clears thread entries too**, and its
  `remove_part=True` form tears down the commentsExtended part alongside
  the comments part.
- **The fabricated `comments.xml` root declares `xmlns:w14` and
  `mc:Ignorable="w14"`**, since comment body paragraphs now carry a
  `w14` attribute.

Documents written by python-docx or by Word before 2013 have no thread
data at all; they read as one unresolved single-comment thread each, and
replying to or resolving such a comment materializes the missing metadata
in place rather than failing.

## [0.3.0] - 2026-06-15

### Added

- **Command-line interface (`docx-plus`)** — a console entry point over
  the library, built on stdlib `argparse` (no new runtime deps).
  `docx-plus inspect FILE` dumps the effective formatting of every
  paragraph (with `--provenance` and `--json`); `docx-plus restyle FILE
  --target ... -o OUT` wraps `remap_styles`; `docx-plus controls
  {list,set,clear}` reads and edits content-control values, coercing the
  command-line string to the control's type (`bool` for checkboxes,
  `datetime` for dates). Read commands take `--json`; mutating commands
  require `-o/--output` (or an explicit `--in-place`) so the input is
  never overwritten by accident. Runnable as `docx-plus` or `python -m
  docx_plus.cli`; documented in `docs/cli.md` and the `cli` agent-skill
  reference. (The agent `SKILL.md` stays repo-level only for now;
  bundling it in the wheel is deferred.)
- **Tracked changes (`revisions/`)** — the OOXML revision marks
  python-docx cannot reach. `mark_insertion` / `mark_deletion` wrap
  existing runs in inline `w:ins` / `w:del` (deletions retag `w:t` to
  `w:delText`); `read_revisions` enumerates every revision type — run-level
  insertions/deletions, move wrappers, run/paragraph property changes, and
  paragraph-mark insertions/deletions — with id, author, timestamp, type,
  and affected text. `accept_revision` / `reject_revision` and the
  `accept_all_revisions` / `reject_all_revisions` bulk forms resolve
  insertions and deletions fully, with safe non-structural transforms for
  move and property-change marks. `enable_track_changes` /
  `disable_track_changes` toggle the document-wide `w:trackChanges` flag in
  `settings.xml`. `RevisionIdRegistry` tracks the single shared revision id
  namespace. A new `docx_plus/examples/track_changes.py` and the
  `revisions` agent-skill reference round out the surface.

## [0.2.1] - 2026-05-21

Post-0.2.0 maintenance: agent-facing docs and a Windows console fix. No
library API changes.

### Added

- **Agent `SKILL.md`** — a repo-level skill manifest so coding agents can
  discover the `docx_plus` surface at a glance, surfaced through a new
  `SKILLS` page in the docs nav (which also restored a link-clean
  `mkdocs build --strict`).

### Fixed

- **Examples under cp1252** — the runnable examples now print ASCII to
  stdout, so `python -m docx_plus.examples.<name>` runs on a default
  Windows console (cp1252) without raising `UnicodeEncodeError`.

### Changed

- **Release docs** — the README and docs index now mark v0.2.0 as
  released.

## [0.2.0] — 2026-05-20

Second cycle. Four new capability modules — anchored comments, layout
extras, bookmarks with cross-references, and footnotes / endnotes —
plus a `core/parts.py` foundation for separate OOXML parts. The release
was extended in-place to close every published "Deferred" bullet and add
a `publishing` module (TOC, captions, Table of Figures), then hardened by
a full pre-publication review whose fixes are recorded below. See
SPEC §15 for the scoped v0.3+ roadmap.

### Added — cascade & API (Session F of issues.md review)

- **M10** — the style cascade now resolves theme **fonts**: `load_theme`
  parses `a:fontScheme`, `ThemeColors` gained a `fonts` map + `font()`
  accessor, and a new `resolve_theme_font(theme, token)` maps a `*Theme`
  token (e.g. `minorHAnsi`) to its concrete typeface (e.g. `Cambria`).
  `resolve_effective_formatting` returns a real `font_name` for
  theme-bound fonts and only flags `partial` when a theme reference
  genuinely fails to resolve.
- **L17** — `clear_all_comments` gained a `remove_part=True` keyword that
  tears down the comments part and its relationship entirely (the default
  still leaves the now-empty part connected for reuse).
- **M11** — `find_matching_style` gained an optional `style_type` filter so
  a wrong-type look-alike (a *character* style named "Heading 1") can no
  longer satisfy a request for the *paragraph* style; `ensure_style` and
  `remap_styles` use it.
- **N4** — `body_document_for(proxy, *, operation=...)` hoisted into
  `core/oxml` (re-exported from `core`) as the shared proxy→`Document`
  resolver for the `comments` and `notes` packages.

### Fixed — correctness (Session F of issues.md review)

- **M3** — the comment-id registry now also seeds from
  `<w:commentRangeEnd>`, so a lone orphaned range-end still blocks id reuse.
- **M6** — `delete_comment` / `clear_all_comments` now remove only the
  `<w:commentReference>` marker and prune its run only when empty, instead
  of dropping a whole `<w:r>` that may carry sibling text.
- **M9 / M13** — `_resolve_color` no longer stores a bare theme name (not
  valid hex) when the theme loaded but the name is unknown, and `partial`
  is set only when a theme reference actually fails — a theme-less document
  with no theme references now resolves `partial=False`.
- **M12** — `delete_style` / `remap_styles` reference scanning now spans
  headers, footers, footnotes, endnotes, and comments parts (not just the
  main body), and `remap_styles` rewrites a target only through the ref tag
  matching the resolved style's type.
- **M5 / M7** — `Border` validates its `color` (ECMA-376 `ST_HexColor`);
  `set_columns` and the mid-document section-break `<w:type>` now use
  schema-strict insertion (`w:cols` before `w:docGrid`; `w:type` after
  header/footer references).
- **M18** — `mark_fields_dirty` and the even/odd-header helpers collapse
  any duplicate `settings.xml` elements instead of acting on only the
  first match.
- **L1 / L15** — comment `w:date` now carries millisecond precision;
  `set_line_numbering` rejects a negative `distance`.

### Changed — internals, docs & tests (Session F of issues.md review)

- **L11 / L13 / L14** — `xpath` caches compiled expressions; the `etree`
  import style is uniform (module-level wherever referenced); `DocxPlusError`
  moved to `core/errors.py`, removing the `# noqa: E402` import ordering.
- **L5 / L6** — `_apply_cell_cascade` dropped its unused `doc` parameter and
  `_classify_target` returns `(kind, element)`, removing three
  `type: ignore[union-attr]`.
- **M20** — the `notes-v0_*` internal-planning cross-references were
  removed from the five capability `__init__` docstrings (and softened in
  CHANGELOG / ARCHITECTURE), so a `pip download` carries no dangling links.
- **M21 / M22 / N6 / N7 / N11** — `conftest` is the single canonical fixture
  path (per-fixture lazy builders into a session tmp dir); `build_fixtures
  main()` is a manual temp-dir helper; the smoke `EXAMPLES` list is derived
  from the package; the LibreOffice render suite now covers all eight
  docx-writing examples.
- **M23** — `docs/TEST_GAPS.md` carries a status note marking its snapshot
  historical (current: 717 tests / 34 files) and flagging the IMPORTANT
  items as the v0.3 re-audit backlog.
- **L2–L4, L8–L10, L16, L21, N2, N5, N8, N12** — docstring / comment
  clarifications, a tidier ToF instruction builder, a real header-paragraph
  fixture replacing an unexercised fake, a more precise frozen-dataclass
  assertion, and reference-page reconciliation (`resolve_theme_font`,
  `body_document_for`).

### Added — style writer parity (Session E of issues.md review)

- **H17** — `create_style` / `modify_style` now accept the six toggle
  properties the cascade resolver already surfaces but the writer
  could not previously produce: `cs_bold` (→ `<w:bCs>`), `cs_italic`
  (→ `<w:iCs>`), `emboss`, `imprint`, `outline`, `shadow`. A
  `ResolvedFormatting` read can now round-trip back through the writer
  for all twelve ECMA-376 17.7.3 toggles instead of hitting
  `UnknownStylePropertyError` on the six new ones. 12 new round-trip
  tests.

### Fixed — docs reconciliation + classifier (Session E of issues.md review)

- **C5** — `pyproject.toml` development-status classifier bumped from
  `3 - Alpha` to `4 - Beta` (conventional for a pre-1.0 surface this
  size). `pyproject.toml` package description now lists `publishing`.
  (PyPI publication banner left as-is pending a publish decision; the
  in-docs `SPEC §…` references are prose, not links — `mkdocs build
  --strict` is clean.)
- **H14** — `docs/ARCHITECTURE.md` §10 test count refreshed (was the
  stale "532"); a new `IMPLEMENTATION.md` §12 progress-log entry
  records the pre-publication review.
- **H15 / M17** — the four exported-but-undocumented exception classes
  (`IdRangeError`, `InvalidNamespaceError`, `InvalidColorError`,
  `InvalidDropdownItemError`) are now documented in
  `docs/ARCHITECTURE.md` §9 and `docs/API.md`. Audited every
  `docs/reference/*.md` `members:` list against its module's
  `__all__`; added the eight v0.2 symbols that had drifted out of the
  rendered reference (edit verbs, `clear_all_comments`, `TableContext`,
  `OffsetFrom`, `build_complex_field`, `insert_before_first_anchor`,
  `XML`, and the new errors).
- **H16** — `SPEC.md` reframed: a status banner marks it the original
  v0.1 design contract and points to `ARCHITECTURE.md` §11 for the live
  roadmap; §15's deferred list is annotated shipped-vs-deferred (§16's
  error table was already current as of Session C).
- **M4** — `edit_comment` / `edit_footnote` / `edit_endnote` `Raises`
  blocks now note that the not-found errors subclass `KeyError`
  (SPEC §16). Also corrected their docstrings (strip "all child
  block-level content", not just paragraphs — matches the H6 fix).
- **M8** — documented that the cascade resolver surfaces only run /
  paragraph properties from table styles; cell / row / table-level
  properties (`<w:tcPr>` / `<w:trPr>` / `<w:tblPr>`) are not resolved
  (deferred to v0.3+). Noted on `resolve_effective_formatting` and
  `TableContext`.
- **M19** — `insert_section_break` `Raises` block now documents the
  second `ValueError` (document has no trailing `<w:sectPr>`).
- **N3** — fixed `\\c` → `\c` in the publishing modules' raw (`r"""`)
  docstrings (double backslash rendered literally).
- **N9 / N10 / N13** — `pyproject.toml` / `mkdocs.yml` descriptions
  mention `publishing`; the README build-phases table collapsed to a
  compact v0.1 / v0.2 summary pointing at `IMPLEMENTATION.md` §12; the
  CHANGELOG's initial-cycle comment / footnote bullets now forward-point
  to the in-place edit verbs added later in this release.

### Added — publishing hardening (Session D of issues.md review)

- **H13** — `add_toc` gained an optional `additional_styles` keyword:
  a sequence of `(style_name, level)` pairs that get appended to the
  TOC via the ECMA-376 17.16.5.61 `\t` switch. Originally listed in
  the v0.2 expansion plan but not implemented in the initial cycle.
- **M15** — `add_caption`'s `label` is now optional; omitting it
  defaults to `f"{caption_type} "` (the universal case). The library
  example now uses the shorter `add_caption(p, caption_type="Figure")`
  form. Pass `""` to suppress the label run explicitly.

### Fixed — publishing input validation (Session D of issues.md review)

- **H11 / M16** — `add_caption(caption_type=)` and
  `add_table_of_figures(caption_type=)` now validate against the SEQ
  identifier rule (ASCII letter/underscore start, then letters /
  digits / underscores). `add_caption(numbering=)` validates against
  the ECMA-376 17.16.4.1 format-picture token set. Each rejection
  raises `ValueError` with a clear message. Closes a real injection
  vector where a malicious `caption_type` like `'Figure" \o "1-9'`
  could inject additional switches into the `TOC \c` instruction.
- **H12** — `add_toc(levels=)` is now validated as a two-int tuple in
  the 1..9 outline range with `lo <= hi`. Reversed, out-of-range,
  wrong-arity, and non-int inputs now raise `ValueError` with a
  clear message at function entry instead of producing silently
  malformed TOCs.
- **M14** — `add_caption`'s docstring now explicitly notes that the
  caption paragraph is *not* automatically restyled to Word's
  built-in `Caption` paragraph style. Auto-applying the style was
  rejected as too opinionated; callers who want it should write
  `paragraph.style = doc.styles["Caption"]`.

New module `docx_plus/publishing/_validate.py` holds the shared
validation helpers (`validate_seq_identifier`, `validate_numbering_picture`,
`validate_outline_levels`, `validate_additional_styles`).

### Fixed — error taxonomy + cascade interleaving (Session C of issues.md review)

- **C4** — SPEC §9.7 and §16 amended to formally bless the raw
  `ValueError` / `TypeError` carve-out for argument-shape validation at
  the public surface, matching what ARCHITECTURE §9 already
  documented. Typed `DocxPlusError` subclasses remain required for
  domain failures (lookup miss, cascade limit, malformed structure,
  etc.). SPEC §16's table now also lists the v0.2-expansion errors
  `CommentNotFoundError` and `NoteNotFoundError`.
- **H9** — `_apply_table_style_chain` rewritten to walk the basedOn
  chain once and interleave base + matching conditional branches per
  style level (ancestors first), per ECMA-376 17.7.6.5. Previously
  the helper applied base for the whole chain then conditional for
  the whole chain — a child style's base could not override an
  ancestor style's matching `<w:tblStylePr>` branch. Helper
  `_apply_conditional_table_formatting` removed (folded into
  `_apply_table_style_chain`).
- **H10** — `protection/document.py` now imports the shared
  `insert_before_first_anchor` from `core.oxml` instead of carrying a
  byte-identical local copy. Eliminates drift risk.
- **M1** — `add_field` now raises `ValueError` on empty or
  whitespace-only `instruction`. Previously emitted a structurally
  invalid field that Word silently rendered as blank.
- **M2** — `add_page_number_field(format="")` and whitespace-only
  `format` are now treated the same as `format=None` (no double-space
  in the emitted instruction). `format` is stripped on the way in.

### Fixed — schema / part wiring (Session B of issues.md review)

- **C1** — Fresh `footnotes.xml` / `endnotes.xml` parts are now seeded
  with the two reserved separator entries (`w:id="-1" w:type="separator"`
  and `w:id="0" w:type="continuationSeparator"`) Word expects per
  ECMA-376 17.11.16 / 17.11.7. Without them, Word may surface
  "needs repair" prompts and strict consumers may reject the file. The
  `read_footnotes` / `read_endnotes` filter already excludes ids ≤ 0,
  so user-visible note iteration is unchanged.
- **C3** — `<w:pgBorders>` child elements are now written in the
  schema-required sequence `top → left → bottom → right` per
  ECMA-376 17.6.10. Previous order was `top, bottom, left, right` —
  permissive consumers accepted it but strict validators rejected.
- **H6** — `edit_comment` and `edit_footnote` / `edit_endnote` now
  strip ALL block-level children before re-appending the new paragraph,
  not just `<w:p>` children. Comments / notes authored elsewhere can
  legally contain `<w:tbl>`, `<w:sdt>`, `<w:customXml>` per
  ECMA-376 17.13.4.2 + EG_BlockLevelElts; the prior filter left those
  siblings next to the new paragraph.
- **H7** — `set_page_borders` now emits `w:offsetFrom="page"` by
  default, matching Word's UI emission. A new `offset_from` keyword
  (`"page"` | `"text"`) lets callers choose. The `Border.space` docstring
  is corrected: the unit is **points** (range 0-31) per ECMA-376
  17.6.10, not twips as previously stated. New `OffsetFrom` literal
  re-exported from `docx_plus.layout`.
- **H8** — `clear_all_comments` is now single-pass O(N+M): one walk over
  the document body removing every range marker / reference regardless
  of id, then one walk over `comments.xml` removing every entry. Prior
  implementation invoked `delete_comment` per comment, repeating the
  full-body scan N times.

### Fixed — cascade correctness (Session A of issues.md review)

- **C2** — Run-level `w:rStyle` now applies *before* direct run rPr per
  ECMA-376 17.3.2.29. Previously the run's own character-style chain ran
  after the direct rPr, so a style-defined property would override a
  direct one. Provenance for run-level `rStyle` is now reported as a new
  `runStyle` layer (distinct from `linkedCharStyle`, which remains the
  paragraph style's `w:link` companion).
- **H1** — Conditional table-style precedence: `_TBL_STYLE_PR_ORDER`
  now lists rows before columns per ECMA-376 17.7.6.5, so at a cell
  matching both `firstRow` and `firstCol` (with no `nwCell` defined)
  the column branch wins, matching Word.
- **H2** — `<w:dstrike>` (double strikethrough) is now read by the
  resolver and surfaced as `ResolvedFormatting.double_strike`. Handled
  as a non-toggle property (last-writer-wins) per ECMA-376 17.7.3 —
  `dstrike` is not in the toggle property list. Independent of `strike`.
- **H4/H5** — Band2 conditional branches (`band2Horz` / `band2Vert`)
  are now reachable: `TableContext` gained `is_band2_row` and
  `is_band2_col` fields, derived as the complement of band1 at the
  default band-size. The resolver now honors
  `<w:tblStyleRowBandSize>` / `<w:tblStyleColBandSize>` when present on
  the table instance's own `<w:tblPr>` (style-chain lookup remains
  deferred — see TableContext docstring).

### Added — initial cycle

- **Anchored comments** (`docx_plus.comments`) — `add_comment`,
  `read_comments`, `delete_comment`, `CommentRef`, `AnchoredComment`,
  `CommentIdRegistry`. Closes the largest python-docx gap: python-docx
  writes the `<w:comment>` body but skips the three body-side anchors
  (`commentRangeStart` / `commentRangeEnd` / the `CommentReference`
  marker run); `add_comment` writes all four, plus creates the comments
  part on first use. Comment threading (w15) deferred to v0.3. In-place
  `edit_comment` / `clear_all_comments` were added later in this release
  — see "Added — in-place expansion" below.
- **Layout extras** (`docx_plus.layout`) — `set_columns` for `<w:cols>`,
  `insert_section_break` for mid-document section breaks (copies the
  trailing `sectPr`'s properties into the chosen paragraph), and
  `enable_distinct_even_odd_headers` / `disable_…` for the doc-level
  `<w:evenAndOddHeaders/>` flag in `settings.xml`.
- **Bookmarks + cross-references** (`docx_plus.bookmarks`) —
  `add_bookmark`, `read_bookmarks`, `delete_bookmark`, plus
  `add_cross_reference` building `REF` / `PAGEREF` complex fields on
  top of `core.build_complex_field`. `BookmarkIdRegistry` lives in its
  own namespace (separate from SDT and comment ids).
- **Footnotes + endnotes** (`docx_plus.notes`) — `add_footnote`,
  `add_endnote`, `read_footnotes`, `read_endnotes`, paired
  `FootnoteIdRegistry` / `EndnoteIdRegistry`. Reserved ids -1 / 0
  (separator / continuationSeparator) are unissuable; `read_*` filters
  separator entries out of results. Insert-only in the initial cycle;
  in-place `edit_footnote` / `edit_endnote` were added later in this
  release — see "Added — in-place expansion" below.
- **`core/parts.py` foundation** — `get_or_create_part(doc, spec)` for
  separate OOXML parts (`comments.xml`, `footnotes.xml`,
  `endnotes.xml`). Registers `XmlPart` subclasses for footnote /
  endnote content types with `PartFactory.part_type_for` so existing
  documents round-trip with parsed XML rather than raw blobs.
- **`core.build_complex_field`** — promoted from `fields/simple.py`'s
  private `_build_complex_field` so cross-references and any future
  field-using module can share it without cross-capability imports.
- **`core.insert_before_first_anchor`** — schema-strict insertion
  helper hoisted from `fields/update.py`. Now used by both
  `fields.mark_fields_dirty` and `layout.enable_distinct_even_odd_headers`.
- **Examples** — `add_comments`, `multi_column_layout`,
  `bookmarks_and_xrefs`, `footnotes_and_endnotes`. Smoke-tested in CI.

### Added — in-place expansion

- **Toggle property completion** — `ResolvedFormatting` now surfaces
  all twelve ECMA-376 17.7.3 toggle properties: the original six
  (`bold`, `italic`, `caps`, `small_caps`, `strike`, `vanish`) plus
  the six complex-script / decorative variants (`cs_bold`,
  `cs_italic`, `emboss`, `imprint`, `outline`, `shadow`). Closes the
  v0.1 "Known limitations" bullet.
- **Comment editing** — `edit_comment(doc, id, text)` and
  `clear_all_comments(doc)`. `CommentNotFoundError` (subclasses
  `DocxPlusError, KeyError`) for missing ids. Body-side anchors and
  `<w:comment>` element attributes (`w:author`, `w:date`,
  `w:initials`) are preserved across edits.
- **Note editing** — `edit_footnote(doc, id, text)` and
  `edit_endnote(doc, id, text)`. `NoteNotFoundError` for missing ids.
  Reserved separator ids (`-1`, `0`) raise `ValueError`.
- **Layout: line numbering** (`docx_plus.layout.set_line_numbering`) —
  emits `<w:lnNumType>` with `count_by` / `restart` / `start` /
  `distance`. Idempotent and schema-strict (sectPr child ordering
  per ECMA-376 17.6.17).
- **Layout: page borders** (`docx_plus.layout.set_page_borders` +
  `Border` dataclass) — emits `<w:pgBorders>` with per-side
  `Border(style, size, color, space)`. All-None removes the element.
- **Conditional table-style formatting** — the cascade resolver
  applies `<w:tblStylePr>` branches (`firstRow`, `lastRow`,
  `firstCol`, `lastCol`, `band1Horz`, `band1Vert`, the four corners,
  `wholeTable`) in ECMA-376 17.7.6.5 precedence order. New
  `TableContext` dataclass; auto-derived from a `_Cell`'s position,
  or pass explicitly to query hypothetical positions.
- **`docx_plus.publishing` module** — `add_toc` (Table of Contents),
  `add_caption` (figure / table captions via `SEQ` complex field),
  `add_table_of_figures` (`TOC \c "Figure"`). Composes existing
  `core.build_complex_field`; users call
  `docx_plus.fields.mark_fields_dirty` before save so Word populates
  results on open.
- **Example** — `publishing_layout` demonstrates TOC + captioned
  figures + ToF. Smoke-tested in CI.

### Quality gates

- `pytest` — 709 passed, 8 skipped (the LibreOffice render tests, gated
  by the `requires_libreoffice` marker).
- `mypy --strict` clean across all modules.
- `ruff check` and `ruff format --check` both clean (Google-convention
  docstrings).
- `mkdocs build --strict` clean.
- Coverage gate at ≥90% holds.
- Examples smoke-tested via `tests/test_examples_smoke.py`.

### Deferred to v0.3+

- w15 threaded comments (parent / child replies, resolve / reopen).
- `STYLEREF` / sequence-field cross-references to headings, captions,
  numbered items.
- CLI (`restyle` + `inspect` + `controls` subcommands).
- Custom XML Parts data binding for content controls.
- Bibliography (sources, citations, `BIBLIOGRAPHY` field) — rides on
  CXML data binding.
- Tracked changes read / write API.
- Glossary placeholder text for SDTs.
- Password-protected forms (legacy hash algorithm).
- See SPEC §15 for the remaining held-beyond items.

## [0.1.0] — 2026-05-19

First public release. The library composes with `python-docx` rather
than replacing it: callers keep their `Document` object and use
`docx_plus` for the operations `python-docx` cannot reach. See
[`SPEC.md`](SPEC.md) for the API contract and the
[docs site](https://thomas-villani.github.io/docx-plus/) for the full
reference.

### Added

- **Style cascade** (`docx_plus.styles`) — `resolve_effective_formatting`
  walks the six OOXML formatting layers (docDefaults, table style,
  paragraph style chain, numbering, direct paragraph, direct run) and
  returns a fully-resolved `ResolvedFormatting` with optional
  per-field provenance. Cycle detection and depth limit (11) enforced.
- **Style modification** — `create_style`, `modify_style`, `apply_style`,
  `delete_style`, `ensure_style`, `list_styles`. Property kwargs share
  field names with `ResolvedFormatting` for round-trip. Schema-strict
  child ordering enforced on `w:style`, `w:pPr`, `w:rPr`.
- **Style remapping** — `find_matching_style`, `remap_styles` for
  reconciling documents whose style ids differ in casing or spacing
  from the canonical Word ids (`"Heading 1"` vs `"Heading1"`).
- **Latent built-ins** — 107 entries in the built-in styles table
  covering Heading 1–9, Title, Subtitle, Quote, TOC 1–9, body / macro
  / preformatted families, comment and footnote/endnote pairs, table
  defaults, and the common character emphasis set. Defaults extracted
  from real Word-saved samples, not guessed.
- **Theme color resolution** (`docx_plus.styles.theme`) — read-only
  parsing of `theme1.xml` and ECMA-376 17.18.40 tint / shade / lumMod
  / lumOff transforms. Missing or malformed themes set
  `ResolvedFormatting.partial=True` rather than raising.
- **Content controls** (`docx_plus.controls`) — `FormBuilder` writes
  text, dropdown, date picker, and checkbox SDTs inline.
  `read_controls` and `set_control_value` round-trip them through
  save/reopen with type-dispatched value handling.
- **Fields** (`docx_plus.fields`) — `add_page_number_field`,
  `add_date_field`, generic `add_field`, and `mark_fields_dirty`
  (sets `w:updateFields` in `settings.xml` so Word recalculates on
  open).
- **Protection** (`docx_plus.protection`) — `protect_document`,
  `unprotect_document`, `is_protected` for `forms` / `readOnly` /
  `comments` / `trackedChanges` modes. Unpassworded (SPEC §1
  non-goal).
- **Typed error hierarchy** — every library-raised error subclasses
  `DocxPlusError`. Errors with builtin analogues (`ValueError`,
  `TypeError`, `KeyError`) multiple-inherit so existing `except`
  clauses still catch them. See SPEC §16 for the full taxonomy.
- **PEP 561 typing marker** — `docx_plus/py.typed` ships so downstream
  `mypy` users see the type hints.
- **Examples** — `docx_plus.examples.inspect_document`,
  `restyle_existing`, `build_form`, `populate_form`. Runnable as
  `python -m docx_plus.examples.<name>` and smoke-tested in CI.

### Quality gates

- `mypy --strict` clean on `docx_plus/`.
- `ruff check` clean with `D` (pydocstyle, Google convention) on the
  library; relaxed on tests and examples.
- Coverage gate enforced at ≥90% on `core/`, `styles/`, `controls/`.
- Layer-3 LibreOffice headless smoke tests gated by the
  `requires_libreoffice` pytest marker; run on the Ubuntu/Python 3.13
  CI job.

### Known limitations

- The cascade resolver surfaces six of the twelve toggle properties
  in `ResolvedFormatting` (`bold`, `italic`, `caps`, `small_caps`,
  `strike`, `vanish`). The other six (`bCs`, `iCs`, `emboss`,
  `imprint`, `outline`, `shadow`) are spec-recognised but not yet
  exposed — extend `_TOGGLE_RPR` and `ResolvedFormatting` in v0.2.
- `set_control_value` for dates renders `"M/d/yyyy"` identically to
  Word; other formats fall back to ISO 8601 until Word re-renders the
  field on next open. The canonical value in
  `w:date/@w:fullDate` is always correct.
- Conditional table-style formatting (`w:tblStylePr` for firstRow /
  lastRow / etc.) is recognised in the cascade walker but deferred —
  the table style's base `pPr`/`rPr` is applied without conditional
  branches. Tracked in [`docs/TEST_GAPS.md`](docs/TEST_GAPS.md) N2.

### Deferred to v0.2

See SPEC §15 for the full list. Highlights: section / header / footer
first-class API, anchored comments, footnotes / endnotes, bookmarks
and cross-references, table cell shading / borders, theme writing,
password-protected forms, content-control binding to Custom XML Parts.

[Unreleased]: https://github.com/thomas-villani/docx-plus/compare/v0.5.0...HEAD
[0.5.0]: https://github.com/thomas-villani/docx-plus/compare/v0.4.0...v0.5.0
[0.4.0]: https://github.com/thomas-villani/docx-plus/compare/v0.3.0...v0.4.0
[0.3.0]: https://github.com/thomas-villani/docx-plus/compare/v0.2.1...v0.3.0
[0.2.1]: https://github.com/thomas-villani/docx-plus/compare/v0.2.0...v0.2.1
[0.2.0]: https://github.com/thomas-villani/docx-plus/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/thomas-villani/docx-plus/releases/tag/v0.1.0
