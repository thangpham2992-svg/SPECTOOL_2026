# SpecWebTool v5.8.70 — NEXIEZ-MRL Version02 Travel / Stops

Source: Sales & Sales Engineering Manual G11L1P, Rev.M, Table 2.1-B/C (2026-05-08).

- MRL Version02: dynamic Stops dropdown 2–22 at 60 m/min, 2–30 at 96/105/120/150 m/min.
- MRL Version02: validate maximum travel against capacity/speed in Table 2.1-B (60/80/105 m); reject unavailable capacity/speed combinations.
- Existing rule for exactly two stops: Travel <= 11,000 mm continues to apply to both models.
- An existing saved invalid Stops entry remains visible to allow Check Spec to pinpoint it.
- Existing DataTable D513/D514/D515 logic is retained, with additional explicit error messages by input.
- MR and its capacity/dimension/decor logic are unchanged from v5.8.68.
- All other features, including Version02 Maximum Decoration Weight, are unchanged.
