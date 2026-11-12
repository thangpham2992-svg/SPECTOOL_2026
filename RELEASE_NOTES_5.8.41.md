# SpecWebTool v5.8.41 — EVRC-C required for special COPs

- NEXIEZ-MR: with COP C760, C770, or C780, selecting any of NSCR-C, AS, SCS-B, or NS-CB requires **EVRC-C**.
- NEXIEZ-MRL Version02: with a COP code starting with M or P (e.g. M1010, P1030E), the same four selections require **EVRC-C**.
- The workbook's **NS-CB** feature is the user's **NSCB**.
- Check Spec shows a named error at the EVRC-C checkbox when the required feature is missing; users can navigate there from the error.
- EVRC-S is not a substitute under this new rule; no automatic enabling or modification of selections.
- For all other COP models, this additional dependency does not apply.
- No changes to the source database, export schema, or existing Excel/PDF values.

Checks: JavaScript syntax validation and isolated rule cases. Browser end-to-end validation not performed.
