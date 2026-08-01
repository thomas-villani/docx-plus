# Linting and the fix plan

`lint/` (v0.6) is the second **composing layer**, sitting where `cli/`
sits: above the capability modules, reading across them, adding no OOXML
knowledge of its own. Every judgement it makes is built on the [cascade
resolver](cascade.md) and the document sweep — which is the point. A
linter that asked Word for an effective value could say *what* a paragraph
looks like; resolving the cascade ourselves is what makes it possible to
say *which layer* set the value, and every consistency rule is an
assertion about a layer rather than about a number.

**Nothing in the layer writes.** `lint` reports; `plan_fixes` produces a
serializable description of a repair. Splitting the report from the
application is deliberate: the fix model gets designed, serialized, and
reviewed at a point where no code path can corrupt a document.

For the calls, see the [linting guide](../guides/linting.md).

## Rule kinds — how an opinionated feature stays in a lean library

The risk a linter carries is that it smuggles a house style into a
library that has spent five cycles refusing to have one. `RuleKind`
is the answer:

- **`consistency`** — the value fights the document's *own* applied
  styles. The **document supplies the target**, so the rule needs no
  configuration and asserts no preference; it says "this deviates from
  what you established elsewhere".
- **`structural`** — an objective defect, true regardless of house
  style: an outline that skips a level, a `REF` to a bookmark that does
  not exist.
- **`policy`** — the value differs from a target the *user* supplied.
  Inert without one, and **no `policy` rule ships enabled**.

That is why the library can report that forty paragraphs resolve
identically under three style ids without claiming this is wrong.

## `Issue` → `Finding`

A rule body yields `Issue`: only what the rule itself knows. The engine
promotes each to a `Finding` by stamping on the id, kind, and severity
from the registration. A rule therefore cannot advertise one severity
in `--list-rules` and emit another, and a rule never restates its own
metadata. An `Issue` *may* override severity for one finding — a rule
whose seriousness depends on what it found is saying something about
that finding, so it outranks even a profile.

Rules receive the **whole swept document**, never one paragraph at a
time: "this font is an outlier", "these two styles resolve identically",
"the outline skips a level" are all comparative, and none is decidable
from a single paragraph. One sweep serves every rule, with provenance
and baselines always on, so cost is paid once however many rules run.

## The fix vocabulary

A `Fix` is a sequence of **named operations** from a closed `FixOp`
vocabulary, not a callable. A plan has to survive being written to a
file, read by a human, and applied by a different process than the one
that built it, and none of that works if an edit is a Python object
holding a bound method.

Naming the public `docx_plus` call each edit would make was the obvious
first design, and it does not work: the central repair — deleting a
direct property from a run — has no public `docx_plus` call. It is
plain python-docx (`run.font.bold = None`). The vocabulary is defined
against its implementation route instead, and property names are
`ResolvedFormatting` field names, so a plan reads in the same terms as
the report that produced it.

`FixSafety` (`safe` / `review` / `destructive`) is about how
*recoverable* an edit is; `Finding.adds_content` is about *what* changes
— content or only formatting. The two are orthogonal, and conflating
them would put "delete a redundant property" and "delete a paragraph"
in the same bucket.

## What only the planner can decide

Three questions are properties of the *set* of findings, so no
individual rule can answer them:

- **Order.** Every operation names a position in the document **as
  swept**, so a deletion partway down invalidates every index below it.
  Deletions therefore sort last and back to front.
- **The content gate.** A fix that removes a paragraph or a style
  definition changes what the document *contains*, not how it looks.
  Those are withheld unless the caller opts in, and reported in
  `deferred` rather than silently dropped.
- **Conflicts.** Two rules can independently claim the same run
  property or overlapping spans of the same text.

The gate runs **before** conflict detection, and the ordering is not
cosmetic: a withheld deletion claims the whole paragraph, so resolving
conflicts first would let an edit that will never be applied veto one
that will.

Claims are per `(run, property)`, per half-open character span, or
`whole` — deliberately finer than per-paragraph. Coarser detection
would make a paragraph carrying several unrelated defects unfixable,
which is the common case rather than an edge one. Text spans are all
measured against the **original** paragraph text, so they are
order-independent and can be tested for overlap without replaying them
— necessary because `plan_fixes` is a pure function of the findings and
never sees the document.

Every finding lands in exactly one of `fixes`, `deferred`, `conflicts`,
or `unfixable`, so a plan accounts for the whole audit rather than
listing only the good news.

## Eleven rules are report-only

Nine of the twenty rules carry a fix. The other eleven do not, and the
list is pinned by a test so a rule cannot quietly gain one. A skipped
outline level can be repaired by promoting this heading or demoting the
one above it, and those produce different documents. Two styles that
resolve identically give no reason to prefer either as the survivor.
Typed indentation needs a number the document does not contain. A plan
that guessed would be the library asserting a house style — exactly
what the rule kinds exist to prevent.

## Profiles

`profile.py` is the one place a house opinion may live: per-rule enable
/ disable and severity overrides, discovered as `docx-plus-lint.json`
beside the document or above it. Precedence is a profile adjusts the
registry defaults, an explicit `select` overrides both, and `exclude`
is applied last and always wins — so naming a rule on the command line
overrides a profile that disabled it. Configuration does not get to
veto a direct question about one document.

A profile may **not** configure a *tag*: "apply this severity to
whatever carries the tag today" is not a stable thing to check in. A
profile naming a rule that does not exist is an error on load rather
than a setting that silently does nothing.

## Not covered

**Applying a plan.** Nothing in v0.6 mutates a document, and the seven
operations have no implementations yet; that is v0.7.

**Anything outside the document body.** Headers, footers, footnotes,
endnotes, and comments are not swept, so a clean lint says nothing
about a document's header.

**A restyle planner.** `duplicate-styles`, `manual-heading-formatting`,
and `font-outliers` all want the same missing capability — "move this
content onto that style" — which is a higher-level operation than the
fix vocabulary models.
