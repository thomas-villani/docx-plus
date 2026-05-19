# `docx_plus.protection.document`

Document-level protection — the thing that turns a document-with-SDTs
into an actual fillable form. `mode="forms"` locks every range outside
of a content control; `"readOnly"` / `"comments"` / `"trackedChanges"`
cover the other Word edit-restriction modes.

Unpassworded in v0.1 (SPEC §1 non-goal). Password-protected forms
(legacy hash algorithm) are deferred to v0.2.

`w:documentProtection` placement follows CT_Settings schema order
(before `w:defaultTabStop`) — see [`ARCHITECTURE.md` §7](../ARCHITECTURE.md#7-fields-and-protection).

::: docx_plus.protection.document
    options:
      members:
        - protect_document
        - unprotect_document
        - is_protected
        - ProtectionMode
