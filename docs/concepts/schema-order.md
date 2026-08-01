# Schema-strict insertion

OOXML containers (`CT_Style`, `CT_PPr`, `CT_RPr`, `CT_Settings`, …) have
**required child ordering**. Inserting an element in the wrong position
produces a file Word will silently "repair" on open — which sometimes
works, sometimes doesn't, and is always a latent bug.

`styles/modify.py` enforces order via three canonical sequences:

- `_STYLE_CHILD_ORDER` (`modify.py:67-90`) — the children of a `w:style`
  element
- `_PPR_CHILD_ORDER` (`modify.py:92-129`) — the children of `w:pPr`
- `_RPR_CHILD_ORDER` (`modify.py:131-...`) — the children of `w:rPr`

Every write goes through `_ordered_insert(parent, new_child, order)`,
which finds the canonical position and inserts there, rather than
appending. The `test_*_children_ordered_correctly` family in
`tests/test_styles_modify.py:277-340` verifies the invariant after
`create_style`. (Verification after `modify_style` is on the test-gap
list — see [`TEST_GAPS.md`](../TEST_GAPS.md) I2.)

All element construction goes through `core/oxml.py`'s `el()` and
`sub()`. No bare `lxml.etree.SubElement` or python-docx `OxmlElement`
calls live in capability modules. This is enforced by the import-invariant
test — see [invariant 2](invariants.md#the-invariants).

## The `settings.xml` case

`w:documentProtection`, `w:updateFields`, and (v0.2) `w:evenAndOddHeaders`
all live deep in `CT_Settings`'s child sequence (ECMA-376 17.15.1.78).
Every callsite applies the same
`core/oxml.py:insert_before_first_anchor(parent, new_element, anchor_tags)`
pattern, walking a tuple of later-siblings (`w:defaultTabStop`, `w:compat`,
`w:rsids`, etc.) and inserting before the first match. If no anchor is
present, the helper falls back to appending — the no-anchor case is
exercised by `test_mark_fields_dirty_appends_when_no_anchor`.

The helper lives in `core/oxml.py` (hoisted in v0.2 when
`layout/settings.py` became the third caller); the per-module anchor tuples
stay co-located with their callsites so the schema position is reviewed
alongside the new child.

`core.ordered_insert` (v0.5) is the stronger form: given the parent's full
child sequence it is idempotent, replacing any same-tag sibling rather than
adding a second. It was promoted out of `styles/modify.py` so
[`numbering`](numbering.md) could share it.
