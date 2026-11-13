# SpecWebTool v5.8.45 — MRL Laminated Validation & Clean Background

- Fixed the MRL Version02 LightUp/conditional-formatting validation mismatch: legacy PPS references are now converted to `Laminated steel sheet` for MRL only.
- Restored valid selection of CP23, CP101, CP141, LA03 and LA02 in WallPattern, CarDoorPattern and CarTransomPattern without erroneous locked-cell Check Spec failures.
- Retained Painted conditional controls and independent MR PPS rules.
- Removed both brochure backgrounds, associated CSS watermark and dynamic model watermark handler.
- Preserved favicon/link previews and existing specification, export, and technical validation logic.

Verification: JavaScript syntax and conditional-formatting test cases (laminate, painted, SUS-HL) passed. No actual browser UI smoke test performed.
