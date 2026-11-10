# QA v3.6

- `node tools/qa_v36.js` — 160 expected OH/PD mappings + 4 exact TR boundaries + 6 Service floor tests: PASS.
- `node tools/regression.js` — 21 existing MR/MRL source-check tests: PASS.
- DataTable: `DA3:DL164`, 160 explicit speed/capacity/TR records; 120 OK, 40 CHECK LAYOUT, source trace `BG[row]:BS[row]`.
- Excel GQXL2P: `SpecForm!F36` = PD minimum, `SpecForm!F37` = OH minimum; validation `DataTable!D523` checks both against the input D36/D37.
- Excel workbook original tables, conditional formatting, and validations retained. Formula engine recalculates in Webtool from new source.
- No graphical browser verification completed in this execution environment (local browser navigation blocked by security policy). JavaScript source evaluator and 21+170 automated cases passed.
