# QA v3.9 — Blue export headings & precise dimensional errors

## Source and scope

Built on SpecWebTool v3.7 (includes AECC A for 3C/4C SAI-2200C, OH/PD range table, free-text Service floor, MRL JJ dropdown); incorporates blue I/II/III/IV headings code from v3.8.
No Excel Database edits: both workbooks in `database/` preserved from v3.7.

## Check Spec messages

- MR: AH/BH use `TblK5_Capacity` result columns `AH-Valid(-Value)` / `BH-Valid(-Value)`; OH uses priority `DataTable!B529` over `E521`; PD uses `E524` over `E523` and model-specific maximum pit limits already coded from legacy Tool.exe.
- MRL: `DataTable!D522` separates AH=SpecForm!D33/F33 and BH=D34/F34; `D523` separates PD=D36/F36 and OH=D37/F37. Range exceeds `TblK5_Capacity[AH-Max]`, `[BH-Max]` or corresponding `-FIRE` columns are flagged by legacy checks.
- Missing lookup -> `CHECK LAYOUT`, not an invented number. Duplicate overall B209 issue suppressed when detailed dimension issue exists. UI error rows provide actual and threshold; clicking navigates to input.
- `SpecForm!B209/B210/B206`, `TblCondition1`, AECC rules and remaining technical checks left intact.

## Automated verification

- `node tools/regression.js`: 21/21 PASS.
- `node tools/qa_aecc_v37.js`: 1C/2C/3C/4C AECC=A; other prohibited features remain prohibited.
- Browser-in-memory loaded all 20 project JavaScript files with no runtime page errors, tested: MR OH/PD/AH/BH lower limit, MR PD/AH/BH upper limit; MRL OH/PD/AH/BH lower limit, MRL AH/BH upper limit; sample error row and click target bindings.
- MR and MRL baseline checks PASS, XLSX exports completed. `xl/styles.xml` and worksheet exported section rows I/II/III/IV/Features all have #D9EDF7 in A:C (5 sections per model); default feature body stays no injected brown fill.
- A4 PDF preview generated from exported XLSX (same user-facing renderer) for MR and MRL. All five Roman headings use colspan 3 and computed rgb(217,237,247). PDFs rendered in Chromium for inspection; no page errors.
- Dynamic MRL 120 mpm without applicable AH/BH lookup reported `CHECK LAYOUT`; MRL fire-door `AH-Max-FIRE/BH-Max-FIRE` bounds verified.

## Caveats

- An arbitrary height above an available minimum is not automatically disallowed unless a maximum is explicitly defined (e.g., MRL AH-Max/BH-Max or MR PD upper range). Never invent OH maxima.
- MRL higher speed / unusual layout combinations may require manual GQ factory layout; source `CHECK LAYOUT` is preserved.
