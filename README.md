# SpecWebTool v5.2 — Feature Matrix / Version02

- Based on Webtool v5.1, preserving MR/MRL Check Spec, ITP, hall buttons, AH/BH, material descriptions, and images.
- Fixes old ALPHA Q applicability flags in the active MRL DataTable using *Sales & Sales Engineering Manual for NEXIEZ-MRL Alpha Package Version2 [G11L1P]*, manual pages 29–38 (PDF pages 1–10), dated 2026-05-08.
- Controls 1C-2BC, 2C-2BC, and SAI-2200C/3C/4C are evaluated separately. A=Available option, S=Standard, X=Not applicable. Inapplicable features stay disabled.
- Adds missing Manual Version02 feature codes as selectable manual-linked options; their selected entries appear in a separate bilingual appendix at the end of Excel/PDF export if they are not originally represented by the source workbook.
- The nine no-charge features FCC-A, FCC-P, FHC-P, SOHS, GCBK, COS, PTC, UPS, DPS auto-select whenever they are applicable, for both MR and MRL. A manual uncheck remains respected; switching to unsupported group clears it.
- Original source Excel file bytes remain in /database/; corrections are applied in the Webtool during workbook load and before checks/exports. The manual does not establish that these features are free of charge; the zero-cost auto-select rule comes from the user's pricing instruction.

## v5.3 Mirror and mass adjustment
- MR: full Table 11-C of AES13-006 Rev AM, 29 cabin codes, for YZ-52A/YZ-53A/YZ-55SN and Not supplied; B750/B1000 not eligible.
- Existing DataTable allowed decoration-weight formula (AA10:AA56) is recalculated with selected mirror type; no rated-capacity reduction. Handrail adjustments remain unchanged.
- Local mirror is an external material; its actual mass must be included in entered decoration weight.
- MRL Version02: mirror dropdown constrained by AA 1000–2200 for YZ-52A/53A; 1100–2100 for YZ-55SN, plus Local.
- Image cards provided for all MR mirror models. The original workbooks are preserved in /database; updates are applied in Webtool at runtime, not in original Excel.
- Mirror selection is checked on its own field, not Capacity.
