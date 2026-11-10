# QA v3.7 — AECC on SAI-2200C

**Source:** *Sales & Sales Engineering Manual for NEXIEZ-MRL Alpha Package Version2 [G11L1P]*, S-1-U4010-E, 2026-05-08, section 6.4.2, *Signal features 1/3*, printed p.35. Row AECC Basic application: 1C-2BC=A, 2C-2BC=A, SAI-2200C=A. Note: AECC cannot be combined with AECH or AIL.

**Database delta:** `DataTable!J366` **X → A** only. Existing `DataTable!K366` and `SpecForm!E165` formulas retained. No MR database change.

**Validation**
- `node tools/qa_aecc_v37.js`: PASS for all four workbook group codes (1C-2BC; 2C-2BC; 3C-SAI-2200C; 4C-SAI-2200C); `DataTable!K366` and `SpecForm!E165` both evaluate A; all four Check Spec PASS; prohibition of other invalid features remains active.
- `node tools/regression.js`: 21/21 PASS.
- Workbook preserved same 63 XLSX ZIP components; sheet layout, table names, formulas and formatting retained.
- The UI has not been manually checked in a GUI browser; test coverage is source formula/Check Spec engine.
