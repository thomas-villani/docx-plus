# docx_plus — Roadmap

The single authoritative roadmap for `docx_plus`. `SPEC.md §15` (a v0.1-era
historical list) and `docs/ARCHITECTURE.md §11` both defer to this file.

`docx_plus` is, and stays, a lean extension to `python-docx` that does the
things `python-docx` can't. Every item below either fills a documented
`python-docx` gap or rounds out a surface already started here. Ideas that
don't fit that charter are routed to sibling projects, not absorbed.

## Current state — v0.6.0 released

Tagged: `v0.1.0`, `v0.2.0`, `v0.2.1`, `v0.3.0` (2026-06-15), `v0.4.0`
(2026-07-26), `v0.5.0` (2026-07-27), `v0.6.0` (2026-07-31). Shipped
capability modules:

| Module | Surface |
|---|---|
| `styles/` | Cascade inspection (`ResolvedFormatting`, all 12 toggles, theme fonts + colors, conditional table styles), modification, remapping; document-wide sweep, `stop_below` baselines, resolved paragraph spacing, unused-style closure (v0.6) |
| `controls/` | Content controls — `FormBuilder`, read / set / clear values |
| `fields/` | Simple + complex fields, `mark_fields_dirty`, `read_fields` (v0.6) |
| `protection/` | `protect_document` |
| `comments/` | Anchored comments — add / edit / delete / clear, over runs, paragraphs, run ranges; threads — reply, resolve / reopen, nested read (v0.4); durable ids + author presence (v0.5) |
| `layout/` | Columns, mid-document section breaks, even/odd headers, line numbering, page borders |
| `bookmarks/` | Bookmarks + `REF` / `PAGEREF` cross-references |
| `notes/` | Footnotes + endnotes — add / edit / read |
| `publishing/` | TOC, captions, table of figures |
| `tables/` | Table / cell borders, table / row / cell shading, merge + unmerge, `w:hMerge` normalization, direct-formatting read (v0.5) |
| `numbering/` | Custom list definitions — define / apply / restart / read; bullet + numbered presets (v0.5) |
| `revisions/` | Tracked changes — mark insertions / deletions, read revisions, accept / reject, track-changes toggle (v0.3) |
| `lint/` | Document linter — 20 rules over the resolved cascade, profiles, and `plan_fixes` (v0.6). Reports and plans; applies nothing |
| `cli/` | `docx-plus` console command — `inspect` (effective formatting), `restyle` (style remapping), `controls` (list / set / clear values) (v0.3), `comments` (list / resolve / reopen threads) (v0.4), `skill` (path / list / show / install the packaged agent skill) (v0.5), `lint` + `plan` (v0.6) |

Suite at the v0.6.0 release: 2055 tests (2043 pass, 12
LibreOffice-skipped), 96% coverage; `mypy --strict`, `ruff`, and
`mkdocs build --strict` all clean.

## v0.6 — shipped: the linter (`lint/`)

The flagship for the cycle, and the first surface here that is a *tool*
rather than a capability. Cleaning up formatting after the content is
settled is the most tedious part of professional `.docx` work, and it is
exactly the job the style cascade resolver was built to reason about.

**v0.6 writes nothing.** The cycle ends at report → plan: findings, and
an inspectable plan describing what a fix *would* change. Applying that
plan is v0.7. This is a deliberate boundary — it forces the fix model to
be designed while nothing can yet corrupt a document, and it means the
whole release is non-mutating.

### Charter note — why this is in scope

Every module shipped so far fills a documented `python-docx` gap:
python-docx cannot write `w:tblBorders`, so `tables/` does. The linter
fills no such gap. It is the first layer here with an *opinion*, which
is a real departure from "lean extension" and is recorded rather than
assumed.

Two things keep it in charter:

- **It is a composing layer, not a capability.** `lint/` sits where
  `cli/` sits — above the capability modules, allowed to import across
  them (SPEC §9.1 forbids that only between siblings). Capabilities
  reach OOXML; composing layers reach capabilities. `lint/` adds no new
  OOXML knowledge of its own.
- **Mechanism ships; opinions are selectable.** The library provides a
  `Rule` protocol, a registry, and a small built-in set. `docx_plus`
  reports that forty paragraphs resolve identically under three style
  ids; it does not assert that this is wrong. Users select and write
  rules.

**No Word / COM dependency.** `wordlive` informs the *ideas* here — it
stays a verification tool for development, never an import. The linter
is pure-Python and cross-platform like the rest of the library.

### Stage 1 — resolver completion — shipped

The linter is the resolver plus clustering, so every hole in
`resolve_effective_formatting` becomes a class of false positives. Two
backlog items turned out to be prerequisites rather than peers, and both
landed here.

**Style-supplied numbering — shipped.** Layer 4 read only the
paragraph's *direct* `w:numPr`, so a correctly-styled `List Bullet`
paragraph reported `num_id=None` — indistinguishable from one where a
bullet glyph was typed by hand, which is exactly the distinction the
`manual-list` rule turns on. It was a live contract bug regardless:
every other field on `ResolvedFormatting` walks the style chain.
Numbering now resolves through `basedOn`, nearest style winning.

Three things worth recording:

- **Reaching the definition at all had a second effect.** Once the
  reference resolves, the `abstractNum` level's own `pPr` applies too,
  so a `List Bullet` paragraph picks up its 360-twip indent for the
  first time. That was not in the plan; it is correct per ECMA-376
  17.7.2 and no existing test moved.
- **`numId` and `ilvl` merge independently**, so a paragraph overriding
  only the level keeps its style's list. The spec states no merge
  semantics for a compound property across the style / direct boundary,
  so this was settled against Word 2016: a `List Bullet` paragraph given
  a bare `<w:ilvl w:val="2"/>` renders as a third-level bullet of the
  style's own list, not as body text.
- **`numId="0"` is surfaced, not flattened to `None`.** It is the
  "explicitly not numbered" sentinel, so `None` now means "no numbering
  information anywhere" and `0` means "deliberately suppressed" — which
  is the difference between a paragraph nobody numbered and one someone
  opted out. New `styleNumbering` provenance layer carries the supplying
  `style_id`.

**Cached document-wide sweep — shipped** as
`styles.iter_resolved_paragraphs`, yielding `ResolvedParagraph` lazily in
document order. Profiling drove the design: `load_theme` alone was 39% of
per-call cost, re-parsing `theme1.xml` once per target. A private
`_ResolverCache` memoizes only what does not vary with the target, giving
**0.611 ms → 0.128 ms per target (4.8x)**.
`resolve_effective_formatting` keeps its exact behaviour and now builds a
throwaway cache, so both paths run one shared walk and cannot drift.

Two walk details a naive implementation gets wrong, both worth keeping in
mind for the rules: `doc.paragraphs` drops tables and `doc.tables` drops
ordering, so document order needs `iter_inner_content`; and `row.cells`
returns a merged cell once per grid column it spans, so a spanned cell's
paragraphs double-count without a `w:tc` dedupe.

Scoped out and documented: headers, footers, footnotes, endnotes, and
comments. Only the main body is swept. `remap_styles` already carries the
part-scanning pattern for when the rules need them.

Deliberately **not** pulled in: the cell-formatting cascade resolver. It
is the largest item on the backlog and it unlocks only the table rules.
Blocking the most valuable rules on the most expensive prerequisite is
the wrong trade; table rules ship in a later wave once it lands.

### Stage 2 — `lint/`, read-only — **shipped**

Engine, registry, CLI, docs, and the whole rule table below. Twenty rules
registered, sixteen default-on; `mixed-run-formatting`,
`stray-empty-paragraph`, `font-outliers`, and `unused-styles` ship off.
`docx-plus lint` takes `--rule` / `--exclude` / `--list-rules` / `--json`
/ `--no-tables`.

Three rules pulled a capability into the layer below them, which is what
the composing-layer split is for — the OOXML knowledge lands in the
capability module and the rule only decides what is worth reporting:

- `unused-styles` → **`styles.find_unused_styles`** and
  **`StyleInfo.is_builtin`**
- `broken-cross-reference` / `caption-manual-numbering` →
  **`fields.read_fields`**, the read half of a module that could only
  write
- `mixed-language` → **`ResolvedFormatting.lang`**

**Running the CLI against a realistic document found two false positives
nothing in the unit tests would have.** `trailing-whitespace` fired on
`"See "` before an unrendered `REF` — a field contributes its *cached
result*, empty in a freshly-written document, so every cross-reference and
page number reported. And `manual-heading-formatting` fired on every
caption, because the template's `Caption` style is bold; the rule now
requires the paragraph be *unstyled*, since anything carrying a
non-default style was styled deliberately whatever it looks like. Both are
the kind of thing only an end-to-end run surfaces.

**`unused-styles` was unusable as first written** and is the sharpest
lesson of the batch. Against a stock `Document()` with two paragraphs it
produced **165 findings** — the entire python-docx style gallery, which is
what the template is *for*. Filtering to author-created styles
(`w:customStyle`, ECMA-376 17.7.4.9) and collapsing unused `w:link` pairs
onto their paragraph half took it to **1**, and that one is the style
someone actually made and never applied. The known-built-ins table was the
obvious filter and the wrong one: it does not cover the table-style
gallery, which is most of what a template ships. A rule that reports
correctly and uselessly is a rule nobody runs twice.

**The baseline detour.** `style-drift` turned out to need the backlog's
"resolve beneath the direct layer" item, so that shipped first as
`stop_below`. Two things fell out of it that were not on any list:

- **A run inside a bold `Heading 1` resolved `bold=False`.** Word writes a
  style and its `w:link` partner with identical `w:rPr`; the resolver
  applied both as independent hierarchy levels, so the ECMA-376 17.7.3
  XOR cancelled every toggle the pair agreed on. The paragraph and the
  run disagreed about the same style. Only visible once the baselines made
  the two answers directly comparable.

  The first fix suppressed the XOR for the `w:link` partner, which made
  the symptom go away on a wrong premise. Driving the resolver against
  live Word afterwards showed the partner is not a cascade layer at all,
  and that the toggle rule diverged from Word in five further ways — the
  resolver agreed on 51 of 74 measured answers. It now agrees on all of
  them; see the `Fixed` entry in `CHANGELOG.md` and the measured table in
  `tests/test_cascade_word_verified.py`. **The lesson worth keeping: for
  anything the format's prose leaves ambiguous, measure Word.**

  Applying that lesson to the rest of the cascade found the same density
  of error one layer over. Conditional table formatting never read
  `<w:tblLook>`, so it painted header rows and banding onto tables that
  had asked for neither; it agreed with Word on 19 of the first 80 cells
  measured. Six causes, now at 1015 of 1015 — see
  `tests/test_tables_word_verified.py`. Two of them are worth
  generalising: **the spec's listed order is not always the application
  order** (Word applies vertical bands over horizontal ones and rows over
  columns, both inverted from 17.7.6.5), and **Word normalises on load**,
  so a `wholeTable` branch that the file plainly contains is discarded
  before anything renders. Neither is visible from the XML alone.

  **Theme colours, third pass.** The flagged item here — "`lumMod` /
  `lumOff` implemented but unwired" — turned out to be a documentation
  bug, not a code one: `w:color` cannot carry those attributes, so there
  is no gap to close. Measuring the layer anyway found two things that
  were real. `<w:clrSchemeMapping>` was ignored, so a document that remaps
  its slots resolved `text1` to black where Word renders it white; and the
  tint/shade transforms ran in floating point, which lands on the wrong
  byte at the integer boundaries these values keep hitting. Exact on 18 of
  47 measured values before, 32 after, with the worst-case error halved to
  one unit. The remaining 15 are enumerated in
  `tests/test_theme_word_verified.py` — Word's rounding at those
  boundaries was not reverse-engineered, and the effort/benefit did not
  justify continuing.

  A method note worth keeping: Word's COM will not tell you a rendered
  theme colour. `Font.Color` returns a theme-*encoded* integer, so the
  ground truth had to come from exporting the document — filtered HTML and
  PDF independently, which agreed, and only then was the fitting worth
  doing.

- **Spacing was wrong twice over, and one of them nobody was looking
  for.** `<w:contextualSpacing>` was ignored outright, which is every list
  paragraph in a stock-template document. Measuring it turned up the
  second: Word does not *add* one paragraph's space-after to the next
  one's space-before, it tops the first up to the second, so a pair sits
  `max(after, before)` apart. The two interact — the top-up is measured
  from the declared space-after even when that space-after was itself
  suppressed. 111 gaps measured, all matching; see
  `tests/test_contextual_spacing_word_verified.py`.

  Method note again: COM was no use here either. `SpaceBefore` reports what
  the cascade declares, which is precisely what `contextualSpacing`
  overrides, so the answer had to come out of Word's layout — PDF export,
  paragraph baselines measured.

  One structural question stayed unmeasured: whether a **continuous
  section break** between two same-style contextual paragraphs breaks the
  suppression. Word turned the probe's break into a page break twice; the
  resolver treats the pair as adjacent, which is what sibling adjacency
  implies.

- **The default paragraph style was never applied at all**, which is most
  paragraphs in most documents — anything without a `w:pStyle` resolved to
  `docDefaults` and reported no style. Going after the cell cascade turned
  this up instead: the cell cascade *was* incomplete, but only as one
  symptom of a missing layer.

  The measurement that mattered was the one that placed it. Sitting the
  default under `docDefaults` matches nearly every reading, and gets one
  case wrong: a `Normal` declaring 20pt against a table style declaring
  36pt renders at 20pt, so the default style beats the table style and has
  to be the paragraph-style layer. Selection is by declaration order (last
  `w:default` wins), measured both ways round, with a fallback to the id
  `Normal`. The two other `w:default="1"` styles turned out to be
  non-events — `DefaultParagraphFont` and `TableNormal` never apply.

  The same probes turned up a second, unrelated divergence: **style
  references were followed without checking `w:type`**. Word ignores a
  `w:rStyle` naming a paragraph style and severs a `w:basedOn` that
  crosses types; the resolver followed all of them. 95 of 96 reads now
  agree — the one is cosmetic, Word's COM naming a `Normal` style in a
  document that has none. See `tests/test_default_styles_word_verified.py`.

  Method note: COM was a perfectly good oracle here, unlike the theme and
  spacing rounds. The question is which style *won*, and `Range.Font` plus
  `Range.Style` answer exactly that. One probe was confounded and had to be
  rebuilt — reusing `List Number`'s `numId` to test style-supplied
  numbering failed because its `abstractNum` carries a `w:styleLink`
  binding it to that style, so even the control went unnumbered.

  Next in line, and still resting on inference: nothing in the run and
  paragraph cascade. `w:tcPr` / `w:trPr` / `w:tblPrEx` cell-level
  properties remain unresolved by design (they carry no `rPr` / `pPr`).
- **Paragraph-mark `rPr` is not a run baseline.** It formats the pilcrow.
  The old paragraph-level baseline folded it in, so a run matching it
  looked redundant when deleting the property would have changed the
  rendering.

Three things the first batch turned up:

- **Toggles need `None` and `False` treated as one value.** An explicit
  `<w:b w:val="0"/>` over an unset toggle renders identically but resolves
  unequal — and that is exactly what select-all-then-clear-formatting
  leaves behind, so it is a case the rule has to catch. Non-toggle
  properties keep `None` distinct, since "inherit" and "Calibri" differ.
- **An excerpt that tidies whitespace hides the defect.** A
  `double-space` finding whose excerpt shows single spaces reads like a
  false positive. Excerpts preserve internal spacing, render tabs
  visibly, and the CLI quotes them.
- **The CLI is a cp1252 surface**, like the runnable examples. A U+2026
  ellipsis in the truncation raised `UnicodeEncodeError` on a default
  Windows console; there are now tests asserting the output encodes.

Design reviewed against `../wordlive/spec-linter.md` (2026-07-27), the
shipped design for the sibling's COM linter. Several of its decisions are
adopted wholesale rather than re-derived — it is a live-validated
catalogue of what actually goes wrong in professional documents.

#### The asymmetry that justifies building this at all

wordlive's own §7c documents the limit docx-plus does not have. Resolving
over COM gives the **effective value plus the applied paragraph style** —
a two-layer compare. It cannot say *which* layer set a property: a
character style via `rStyle`, a numbering level's `rPr`, a table-style
conditional branch, and docDefaults are all invisible behind one number.
And Word exposes no per-property "reset to style", so a fix either writes
the style's value back **as a redundant direct property** — leaving the
mess it set out to clean — or calls `Font.Reset()` and destroys
intentional formatting along with the drift.

`resolve_effective_formatting` already returns per-field
`FormattingSource` (layer, `style_id`, `chain_depth`, toggle-resolved),
and OOXML lets us **delete the offending `w:b` / `w:sz` outright**. So the
regularizer here is strictly more precise at the core operation, not the
same thing minus Word. wordlive's spec names our provenance as the
upgrade it wants; that is the differentiator to build around.

#### Adopted from the wordlive design

- **Three rule *kinds*.** `consistency` (no config — a direct override
  deviating from the applied style), `structural` (an objective defect),
  `policy` (needs a profile supplying a target). This is a better answer
  to "how do opinions stay out of a lean library" than the flat rule list
  above: consistency and structural rules judge a document against
  *itself*, and anything requiring a house opinion is inert without a
  profile.
- **`Finding` shape** — rule id, kind, severity, locator, message,
  `fixable`, `fix`, `observed` / `expected`.
- **The `fix` is a call into our own public API**, mirroring wordlive's
  "`fix.op` is literally an exec op". Fixes route through
  `styles.modify_style`, `numbering.apply_list`, and friends rather than a
  parallel writer, so the fix path stays on the audited, tested surface.
- **`default_on` per rule, plus tags** so a user enables a cluster
  (`--rules typography`) instead of naming ids. Unambiguous defects ship
  on; heuristic or opinion-flavoured rules ship off.
- **`adds_content` gate** — a fix that inserts or deletes content is
  withheld unless explicitly allowed. (v0.7, but the flag is designed in
  now.)
- **Idempotency as a test invariant** — regularize twice, assert the
  second pass applies nothing.

#### Rules

All paragraph / run / style scoped, so none waits on the cell cascade.
Starred rules come from the wordlive catalogue.

| Rule | Kind | What it catches |
|---|---|---|
| `redundant-direct-formatting` | consistency | a run's direct `rPr` setting a property to the value it already inherits |
| `style-drift`* | consistency | a direct override deviating from the applied style — wordlive's central rule, and where our provenance beats its two-layer compare |
| `duplicate-styles` | consistency | two or more style ids resolving to identical formatting (`find_matching_style` is the seed) |
| `unused-styles` | structural | defined, referenced nowhere (`remap_styles` / `delete_style` already scan references) |
| `manual-heading-formatting`* | structural | a bold / large `Normal` paragraph that looks like a heading but is not styled |
| `heading-level-skip`* | structural | the outline jumps H1 → H3. Newly cheap: `outline_level` resolves through the style chain after stage 1 |
| `empty-heading`* | structural | a heading paragraph with no text |
| `manual-list` | structural | list-like literal text (`1.`, `a)`, `•`) with no `numPr` |
| `list-numbering-continuity`* | structural | a contiguous run of numbered paragraphs split into independent lists — the "N separate 1. lists" footgun, reachable now that numbering resolves |
| `direct-numbering-override` | consistency | a paragraph's direct `numPr` fighting the one its style supplies |
| `trailing-whitespace`* | structural | a paragraph ending in space / tab |
| `double-space`* | consistency | runs of 2+ spaces in body text |
| `space-before-punctuation`* | consistency | ` ,` ` .` ` ;` ` :` |
| `indent-by-whitespace`* | structural | leading tabs / spaces standing in for a real indent (wordlive's `leading-whitespace`) |
| `stray-empty-paragraph`* | structural | empty paragraphs standing in for `spaceAfter` |
| `font-outliers` | consistency | thinly-populated effective font / size combinations against a dominant set |
| `mixed-language` | consistency | inconsistent `w:lang`, which quietly wrecks spellcheck |
| `broken-cross-reference`* | structural | a `REF` / `PAGEREF` naming a bookmark that does not exist — `bookmarks/` already reads both sides |
| `caption-manual-numbering`* | structural | a `Caption` paragraph numbered with literal text, not a `SEQ` field |

Deliberately out: everything requiring **pagination or a live spell
check**. `table-repeat-header` needs to know a table crosses a page,
`paragraph-too-long` needs page geometry, `toc-present-and-current` needs
field results, and proofing needs Word's dictionary. Those stay
wordlive's. Presence-only halves (is there a TOC at all, does the footer
carry a `PAGE` field) are reachable offline and can come later.

Worth noting the reverse also holds: `mixed-run-format` is report-only in
wordlive because COM returns `wdUndefined` for a paragraph whose runs
disagree and "which run is the outlier needs a run-walk". The stage-1
sweep hands us that run-walk for free, so the rule is *more* capable
here.

### Stage 3 — report → plan — **shipped**

`plan_fixes(findings)` returns a `FixPlan`, and nothing applies it.
Findings carry a `Fix`; `fixable` is now derived from it rather than
stored, so the flag cannot drift from the repair.

**The fix vocabulary is closed, not callable.** Seven named operations
(`FixOp`) with JSON arguments, rather than a bound method per rule. A
plan has to survive being written to a file, reviewed, and handed to a
different process than the one that built it, and none of that works if
an edit is a Python object. The plan serializes end to end.

Three things fell out of building it that the design note above did not
anticipate:

- **The planner's real job is the three decisions no rule can make**,
  because each is a property of the *set* of findings. Order (deletions
  last and back to front, since every operation names a position in the
  document as swept). The content gate. Conflicts.
- **Conflict detection has to be finer than "the same run".** Claims are
  per property and per half-open character span, so two rules clearing
  two different properties of one run both apply, and a `double-space`
  span adjacent to a `space-before-punctuation` span composes. Coarser
  detection would call a paragraph carrying several unrelated defects
  unfixable, which is the common case rather than an edge one. The gate
  runs *before* conflict detection, so a withheld deletion cannot knock
  out an edit that is actually going to happen.
- **Text fixes have to be spans against the original text.** Anything
  phrased as "find this, replace it" cannot be checked for overlap
  without replaying it, and `plan_fixes` never sees the document. Writing
  the fix that way exposed a detection bug in passing: `double-space`
  matched `\S {2,}\S`, which consumes the word between two double spaces,
  so `"a  b  c"` reported one occurrence and would have been half-fixed.

**Nine of twenty rules carry a fix, and the eleven that do not are the
finding.** A skipped outline level can be repaired by promoting this
heading or demoting the one above it, and those produce different
documents; two styles that resolve identically give no reason to prefer
either as the survivor; typed indentation needs a number the document
does not contain. Each rule's docstring says which case it is, and the
list is pinned by a test — a rule quietly gaining a fix is a decision
about someone's document. `direct-numbering-override` is the sharpest
of them: it carries a fix *except* for `numId=0`, which is the ECMA-376
opt-out sentinel and the one override somebody meant.

The **high-level "restyle" planner** on the backlog is still the missing
engine for the style-related rules — `duplicate-styles`,
`manual-heading-formatting`, and `font-outliers` are all report-only for
want of it, since each needs "which style should this be" answered.

**Profiles shipped** as designed: per-rule `enabled` / `severity` /
`options`, loadable from a path, a mapping, or nothing, and discovered as
`docx-plus-lint.json` beside the document or above it. `options` is read
by nothing, because no policy rule ships. One rule settled while writing
it: **an explicit `--rule` beats a profile that disabled it**, so a
checked-in file can never stop someone asking a direct question of one
document. A profile may not configure a *tag* — "apply this severity to
whatever carries the tag today" is not a stable thing to check in — and a
profile naming an unknown rule fails on load rather than silently doing
nothing.

### CLI — shipped

`docx-plus lint FILE` and `docx-plus plan FILE`, sharing `--rule` /
`--exclude` / `--no-tables` / `--profile` / `--no-profile`; `plan` adds
`--allow-content`. Both are read commands, so both take `--json` and
neither needs `-o/--output` — the mutating-command convention is
untouched this cycle. `plan` exits `1` when the plan holds any edit,
applied or withheld, so it gates a pipeline on "is there anything a
repair pass would do"; findings nobody can repair do not fail that gate.

### Docs — shipped

The public-facing surface for the cycle, closing the two gaps stage 3
left open and three more found while writing them:

- `skill/reference/lint.md` — the agent-facing page for the flagship
  feature, with the rule catalogue, the fix vocabulary, profiles, and
  how to write a rule. `skill/reference/cli.md` gained `lint` and `plan`
  sections, and `SKILL.md` its capability-map row.
- `docs/API.md` gained a `docx_plus.lint` section covering all
  twenty-four public symbols, plus the two new CLI commands.
- **`ARCHITECTURE.md` had no `lint/` section at all** — §7.15 now
  carries the design rationale (why a composing layer, why `Issue` →
  `Finding`, why a closed op vocabulary rather than named public calls,
  why claims are per-property, why the gate runs before conflicts), and
  the source tree listing gained `lint/`, the two CLI modules, and the
  four examples it had been missing since v0.3.
- **`README.md` and `docs/index.md` still listed "a document linter" as
  *backlog*.** Both now carry a Lint capability row instead.
- `docx_plus/examples/lint_document.py` — every other capability ships
  a runnable example; this one did not.

The version re-stamp (`API.md`, `SKILLS.md`, README, `docs/index.md`)
was done ahead of the bump rather than after it, which is the first
cycle that has not left those four lagging a release behind.

### Review — done, and what it deferred

Before the release, three reviewers went over everything since v0.5.0
(82 files, ~16.5k insertions): the `lint/` package, the `styles/`
rewrite, and the CLI / skill / docs / packaging surface. Five confirmed
critical bugs and around fifteen significant ones were fixed; all are in
the CHANGELOG. Two findings are worth recording as *deliberately* left:

- **`resolve_paragraph_spacing` builds a throwaway `_ResolverCache` per
  call**, so it is ~6x slower per paragraph than going through the sweep
  (2144 vs 363 microseconds, measured over 300 body paragraphs). That is
  the same shape as `resolve_effective_formatting`, which is documented
  to work that way, so it is consistent rather than wrong — but there is
  no batch entry point for spacing. A `include_spacing=` flag on
  `iter_resolved_paragraphs` is the obvious answer, and it is v0.7 work.
- **`modify.py` open-codes the ST_OnOff test twice more.** Both use the
  correct value set, so this is duplication rather than the bug fixed in
  the resolver; unifying them means either exporting `_on_off` across
  modules or promoting it into `core/`, and neither is worth doing
  mid-release.

The tail of consistency items the review raised — an exhaustive `match`
over `FixOp` so a new operation cannot silently lose conflict detection,
`Literal` types on `_Claim`'s string fields, `Profile.option` reaching
the four hard-coded rule thresholds, and the shared "dominant value"
helper the rule modules each reimplement — are all real and all
non-blocking.

## v0.7 — sketched: the regularizer

Applies a `FixPlan`. Each rule that is safely invertible gains an
autofix; `docx-plus regularize FILE -o OUT` re-enters the
mutating-command convention, dry-run by default. Sequenced after a
release of real-world lint output, because detection is safe and
rewriting someone's formatting is where the risk lives.

## v0.5 — shipped

### Agent skill packaging — shipped

The decision the v0.3 CLI unblocked and v0.4 deferred. The skill moved
from repo-level `skills/docx-plus/` into `docx_plus/skill/`, so it ships
in the wheel, and a `docx-plus skill` command
(`path` / `list` / `show` / `install`) puts it where an agent will find
it.

Two things worth recording:

- **The move needed no build configuration.** Hatchling's
  `packages = ["docx_plus"]` already sweeps non-`.py` files — the reason
  `py.typed` ships — and the sdist `include` already listed
  `docx_plus/`. Confirmed by building a wheel, unzipping it (all ten
  Markdown files present), installing into a clean venv with no source
  tree, and driving `skill path` / `list` / `install` from there.
- **`docs/SKILLS.md` was making a false claim.** It said the library
  "ships" the skill while linking exclusively to GitHub blob URLs, which
  is broken for anyone who `pip install`ed. That is now true rather than
  aspirational, and the page documents the CLI instead of a clone.

Its table had also drifted — missing the `numbering` and `tables`
topics added earlier in this cycle. The suite now asserts every topic
file is linked from `SKILL.md`, so a new page cannot land orphaned.

### Comment durable ids + author presence — shipped

The two side-parts split out of the v0.4 cycle. Both were verified
against a file **Word 2016 authored itself** — driven over COM, saved,
unzipped, and read — before a line was written, which was the right call
twice over:

- **All four content-type / relationship URIs guessed in the groundwork
  PR turned out exact.** No corrections needed.
- **`w16cid:durableId` is hex, not decimal.** The plan specified a
  decimal collector and a decimal registry. Word writes
  `ST_LongHexNumber` — the same 8-uppercase-digit form as
  `w14:paraId` — so `DurableIdRegistry` reuses the existing hex
  machinery and the feature added *less* core code than budgeted.
- **`w15:person` is not a bare element.** It carries a
  `<w15:presenceInfo>` child with `providerId` and `userId`.

Durable ids are written automatically, because they are a comment's only
identifier stable across edits and Word regenerates missing ones anyway.
Author presence is **opt-in** (`set_author_presence`): it is cosmetic,
and registering an author means inventing a `userId` for someone the
library knows nothing about.

Round-trip verified: Word opened a document written here and resaved it
preserving both `paraId` and `durableId` byte-for-byte, plus a
non-default `providerId="AD"` presence entry, and added no parts of its
own.

Deferred: `commentsExtensible.xml` — see the backlog.

### Table borders / shading / merging — shipped

Landed as a new `tables/` capability. The backlog entry bundled three
things at very different states, which is worth recording:

- **Borders and shading were 100% greenfield.** python-docx has no
  `CT_Border`, `CT_TblBorders`, `CT_TcBorders`, or `CT_Shd` class and
  registers none of those tags. Structurally they are `set_page_borders`
  again, over the `Border` dataclass promoted into `core/` in the
  groundwork PR.
- **Merging already worked.** `_Cell.merge` is fully implemented,
  including the non-rectangular check. `merge_cells` only translates its
  `InvalidSpanError` into a `DocxPlusError` subclass. What was genuinely
  missing is the *inverse*: nothing in python-docx removes a
  `w:gridSpan` or a `w:vMerge`, so a merge was one-way. That is
  `unmerge_cell`.
- **`w:hMerge` was a third thing entirely.** OOXML has two horizontal
  merge encodings and python-docx models only `w:gridSpan`, so a table
  written with `w:hMerge` reads back as separate cells Word draws as
  one. `normalize_horizontal_merges` converts between them.

Verified against Word 2016: the example opens with no repair prompt and
renders the banner span, header shading, inside rules, and the
double-rule callout as intended; an `hMerge` fixture and its normalized
form rasterise byte-for-byte identically. Word's own COM object model
turned out to share python-docx's blind spot, reporting six cells for
the `hMerge` fixture and five after conversion while laying both out the
same way.

Deferred: the **cell-formatting cascade** (table style →
`w:tblStylePr` conditional branch → direct `w:tcPr`).
`read_table_formatting` reports direct formatting only, so a
`Table Grid` table reads back with no borders. That resolver is larger
than every writer in this cycle put together — see the backlog.

### Cross-references to non-bookmark targets — shipped

The backlog called this "mostly instruction grammar over existing
plumbing". That was true for one half and wrong for the other, which is
worth recording:

- **`STYLEREF` was instruction grammar** — it already worked through
  `add_field`. What landed is the typed wrapper
  `fields.add_style_reference`, style-name validation, the switch
  surface, and an outline-level shorthand.
- **Caption references were not.** A `REF` field cannot point at a `SEQ`
  field, only at a bookmark, and `add_caption` created none — so "see
  Figure 3" was not expressible at all, with or without new grammar. The
  fix is `bookmark_name` on `add_caption`, which required promoting
  bookmark emission into `core` (SPEC §9.1 blocks `publishing` from
  importing `bookmarks`) and a bookmark *name* registry, which did not
  exist.

Also filled in the rest of `add_cross_reference`'s switches (`\n` `\r`
`\w` `\p` `\t` `\#` `MERGEFORMAT`), which had only ever supported `\h`,
and added the bookmark-name validation it was missing.

Verified against Word 2016: `REF` to a caption bookmark resolves to
`Figure 1`, `PAGEREF` to `1`, `\p` to `above`, and the header `STYLEREF`
renders `Architecture` on page 1 and `Operations` on page 2.
### Custom numbering — shipped

The largest remaining `python-docx` gap. It exposes a `NumberingPart`
and `len()` of its definitions; there is no `CT_AbstractNum` and no
`CT_Lvl`, so nothing in it can say what a list *looks like*. Landed in
`numbering/`.

Shipped:

- **Define** — `define_list_definition` over `LevelDefinition`, plus
  `define_bullet_list` / `define_numbered_list` presets using Word's own
  glyph and format cycles.
- **Apply** — `apply_list` / `remove_list`, the latter able to write the
  `numId="0"` sentinel that suppresses numbering a *style* applies.
- **Restart** — `restart_list`. OOXML has no paragraph-level "count from
  1 again"; this adds a second `<w:num>` over the same
  `<w:abstractNum>` with a `<w:startOverride>`, as Word does.
- **Read** — `read_list_definitions` returning `ListDefinition` /
  `ListLevel`, tolerant of dangling references and malformed ids.

New plumbing: `NUMBERING_SPEC` (needed because
`DocumentPart.numbering_part` fabricates through an unimplemented stub
that raises), two sequential-allocation registries, and
`assert_numbering_well_formed`.

Verified against Word 2016 via `wordlive` — which caught a cramped
outline where the hanging indent was narrower than the rendered
`1.1.1.`, collapsing the separator tab. Now documented on
`LevelDefinition.hanging`.

Deferred to the backlog: linking a definition into a style definition.

## v0.4 — shipped

### Threaded comments — shipped

The reply / resolve model Word has used since 2013, and the second half
of the collaboration story `revisions/` opened in v0.3. Landed in
`comments/threads.py` over a new `comments/_extended.py`.

Shipped:

- **Replies** — `reply_to_comment` parents a new comment to an existing
  one via `w15:paraIdParent` and mirrors the parent's body-side anchor
  range, which is what makes Word render a thread as a single balloon.
- **Resolve / reopen** — `resolve_comment` / `reopen_comment` toggle
  `w15:done` across the whole thread, matching Word's thread-wide
  Resolve button.
- **Nested read** — `read_threads` returns `CommentThread` (root,
  replies, resolved); `read_comments` results gained `parent_id` and
  `resolved`.
- **Eager metadata** — `add_comment` now stamps `w14:paraId` and writes
  an unresolved `<w15:commentEx>` entry, so every comment is
  thread-ready. `edit_comment` preserves that stamp; `delete_comment`
  cascades to replies by default.
- **CLI** — `docx-plus comments list / resolve / reopen`.

New plumbing: `COMMENTS_EXTENDED_SPEC` (fourth separate part),
`ParaIdRegistry` (the first *package*-wide id namespace), the `w15`
namespace, and a `BUILD_NSMAP` / `NSMAP` split so the extension prefix
stays out of `document.xml`.

Deferred to the backlog: `commentsIds.xml` (`w16cid` durable ids, which
Word regenerates) and `people.xml` (`w15:people` author presence,
cosmetic only).

## v0.3 — shipped

Delivered in the v0.3 cycle.

### 1. Tracked changes (read/write) — shipped (v0.3)

Read/write API for OOXML revision marks, landed in the `revisions/`
module. The canonical "`python-docx` can't do this" gap.

Shipped:

- **Read** — `read_revisions` enumerates every revision type (`w:ins`,
  `w:del`, move wrappers, `w:rPrChange` / `w:pPrChange`, paragraph-mark
  insertions / deletions) with id, author, timestamp, type, and text.
- **Write** — `mark_insertion` / `mark_deletion` wrap existing runs;
  `enable_track_changes` / `disable_track_changes` toggle the
  `settings.xml` flag.
- **Accept / reject** — `accept_revision` / `reject_revision` and the
  `accept_all` / `reject_all` bulk forms resolve insertions and deletions
  fully, with safe non-structural transforms for move and property-change
  marks.

Reused the range/target-normalization pattern from `comments/`, the
`_IdRegistryBase` subclass pattern, and the `settings.xml`-touch pattern
from `fields/update.py`. (Revision marks live inline in `document.xml`, so
no separate part was needed.)

Deferred to the backlog: authoring move pairs and property-change markers
(both need a diff engine), and true paragraph merge/split on accept/reject
of paragraph-mark revisions (currently a non-corrupting fallback).

### 2. CLI — `docx-plus` — shipped (v0.3)

A command-line surface over the existing library, landed in the `cli/`
module and exposed via the `docx-plus` console entry point (also
runnable as `python -m docx_plus.cli`). Built on stdlib `argparse`, so
no new runtime dependency.

Shipped:

- `inspect` — dump effective formatting per paragraph (wraps the cascade
  resolver); `--provenance` and `--json`.
- `restyle` — style remapping (wraps `styles.remap_styles`);
  `--target` / `--map` / `--create-missing`.
- `controls` — `list` / `set` / `clear` content-control values, coercing
  the command-line string to the control's type. `set` / `clear` target by
  `--tag` or `--control-id`, since Word writes most controls with an empty
  `w:tag`.

Read commands take `--json`; mutating commands require `-o/--output`
(or an explicit `--in-place`) so the input is never overwritten by
accident.

The deferred **packaging decision for the agent `SKILL.md`** this
unblocked was taken in v0.5: the skill moved into the package and got a
`docx-plus skill` subcommand.

## Backlog — bounded, unscheduled

Each reuses existing plumbing; pull into a cycle as priority dictates.

- **`commentsExtensible.xml`** — a *fifth* comment side-part
  (`w16cex`, 2018), discovered while verifying the v0.5 work against a
  Word-authored file. Keys off `w16cid:durableId` and carries
  `w16cex:dateUtc`, which exists because `w:comment/@w:date` is local
  time. Low priority: `commentsIds.xml` predates it by two years, so
  writing one without the other is a state Word itself produced for
  years — and Word added no such part when resaving a file of ours that
  lacked it. (`comments/`.)
- **Glossary placeholder text** — the "formal" placeholder mechanism for
  SDTs, vs. the inline `w:placeholder` text `controls/` already supports.
- **Password-protected forms** — legacy hash algorithm, paired with
  `protect_document`. (`protection/`.)
- ~~**Resolve beneath the direct layer**~~ — **shipped in v0.6** as
  `resolve_effective_formatting(..., stop_below=Layer)`, with
  `iter_resolved_paragraphs(..., include_baseline=True)` doing it for a
  whole document against the shared cache. It unblocked `style-drift`,
  freed `redundant-direct-formatting` to check `w:rStyle` runs instead of
  skipping them, and surfaced a resolver bug of its own (see the v0.6
  notes). `direct-numbering-override` is now writable via
  `stop_below="numbering"`. (`styles/`.)
- **Sweep the non-body parts** — `iter_resolved_paragraphs` covers the
  main document body only, so any lint rule over headers, footers,
  footnotes, endnotes, or comments has a blind spot there. Confirmed as
  acceptable for v0.6 (the body is the concern), and bounded when it
  comes: `remap_styles` / `delete_style` already carry the
  reference-scanning pattern across exactly those five parts, so the walk
  is a port rather than a design. (`styles/`.)

## Backlog — larger or dependency-gated

- **Custom XML Parts data binding** — wires repeating-section content
  controls to a custom XML data source: new relationship types and
  `<w:dataBinding>` children on SDTs. `core/parts.py` already supports
  separate parts. **Gates the next item.**
- **Bibliography** — sources in a Custom XML Part, `<w:sdt>` citations
  referencing them, a `BIBLIOGRAPHY` field rendering the list. Rides on
  the data-binding subsystem above.
- **Theme writing** — `styles/theme.py` reads themes today; writing rounds
  out the surface.
- **High-level "restyle" planner** — inverse of the inspector: take a
  target `ResolvedFormatting` and compute the minimal cascade modification
  to reach it. Large design space. **Scheduled into v0.6 stage 3**, where
  it is the fix engine behind the style-related lint rules; designing it
  against a plan that cannot yet execute is the cheap place to do it.
- **Sections / headers / footers first-class API** — wraps the
  `python-docx` primitives behind a `docx_plus`-native surface
  (`sections/`).
- **Cell-formatting cascade resolver** — resolve a cell's effective
  borders, shading, and margins through table style → `w:tblStylePr`
  conditional branch (first row / last column / banding) → direct
  `w:tcPr`. `tables/read.py` reports direct formatting only, and
  `styles/inspect.py` scopes this out in the same terms while resolving
  the paragraph and run cascade. The largest single item on this list.
  **Gates the linter's table rules**, and was deliberately kept out of
  v0.6 for that reason — see the v0.6 stage 1 note.
- **Link a numbering definition into a style** —
  `w:style/w:pPr/w:numPr`, so `ensure_style("ListBullet")` produces a
  style that actually bullets. Split out of the v0.5 numbering cycle
  because `styles/modify.py` already owns writing into `w:style` and
  carries `_STYLE_CHILD_ORDER` / `_PPR_CHILD_ORDER`; duplicating that in
  `numbering/` would put the same schema knowledge in two places.
  (`styles/`.)

  Concretely, so this doesn't need re-deriving:

  - The seam is documented in the code already —
    `styles/modify.py` `_BUILTIN_STYLES` Tier D carries a comment saying
    the `List*` styles omit `numPr` and that "callers wanting actual
    auto-numbering should attach a numbering definition separately".
    Asymmetry to fix: the *bundled template's* `ListBullet` does link
    `numId` 1, so `ensure_style` on a document lacking the style
    produces something weaker than a stock `Document()` already has.
  - The work is a new key in `_write_paragraph_property`
    (`modify.py:835`) plus an entry in `_validate_property_keys`
    (`:818`), which currently rejects anything outside the
    paragraph/run property sets. `_PPR_CHILD_ORDER` already reserves the
    `numPr` slot; `_ordered_insert` places it.
  - No cross-capability import is needed: linking is just writing
    `w:numPr` with an int, so `styles/` never has to reach into
    `numbering/`. That is why this lands in `styles/`, not here.
- **Resolve style-supplied numbering in the cascade** — related, and
  arguably the more surprising half. **Scheduled into v0.6 stage 1** as a
  linter prerequisite; the derivation below is the implementation note.
  Layer 4 reads only the paragraph's
  *direct* `w:numPr` (`styles/inspect.py:428`), never the one its style
  supplies, so on a stock `Document()`:

  ```python
  p = doc.add_paragraph("bulleted", style="List Bullet")
  resolve_effective_formatting(p).num_id   # -> None
  ```

  even though that style links `numId` 1 in the bundled template
  (verified 2026-07-26). Every other property on `ResolvedFormatting`
  walks the style chain, so `num_id` silently breaks the contract the
  rest of the dataclass sets. Fixing it means resolving `numPr` through
  the `basedOn` chain like the other pPr properties, and deciding what
  `numId` `0` from a style should mean (the "no numbering" sentinel).
  Wants its own tests; do it alongside or before the style-linking item
  above, since that item makes the gap much easier to hit. (`styles/`.)

## Considered, not on the roadmap

- **MCP server** — surfaced in the old `notes.md` scratchpad alongside the
  CLI. An MCP wrapper is an adjacent product, not a `python-docx`
  extension; route to a sibling project rather than absorbing it here.
  Revisit only if the CLI lands and a thin MCP front end over it is
  cheap.
