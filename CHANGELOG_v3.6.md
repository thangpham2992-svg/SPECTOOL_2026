# v3.6 — OH/PD normalized source lookup and optional Service floor

## GQXL2P / NEXIEZ-MRL Version02 ALPHA

- The original manufacturer's BG8:BS43 table is retained unchanged. New normalized mapping in `DataTable!DA3:DL164`: speed × capacity × travel range, with OH Mitsubishi, OH EN81, PD and SETS-L fields, status and original row citation.
- Travel bands: (0,30], (30,60], (60,80], (80,105] metres. Exactly one range is selected, including at boundaries 30/60/80/105 metres.
- `SpecForm!F36` (minimum PD) and `SpecForm!F37` (minimum OH) now calculate directly from DA:DL. Unsupported combinations yield `CHECK LAYOUT`; they are never inferred from another speed, capacity or range.
- UI reads those same two workbook formulas. When Capacity, Speed, Travel or Regulation changes, any below-minimum OH/PD is raised to minimum; larger manually entered values are not reduced. SETS-L feature also triggers recalculation.
- Check Spec locates independent PD and OH errors at the corresponding inputs.

## Service floor (MR and MRL)

- Service floor text is fully optional and free-format; no automatic floor-name parsing or Stops-count matching. Empty or descriptive strings are accepted by Check Spec.
- Stops and opens remains independently checked for its supported technical range.

## Source / quality

- 160 matrix scenarios in `tools/qa_v36.js`: 120 matched source rows and 40 explicitly unsupported (including speed 60 >60 m and 630/1800 kg at 120/150 m/min). All match normalized DataTable source. 4 exact TR boundaries and 6 Service floor trials pass.
- 21 existing regression tests pass. Source workbook original 40 Excel tables, 50 Data Validations and 22 Conditional Formatting blocks are retained.
- Physical engineering applicability and manufacturer approval of each OH/PD value remain subject to verifying the manufacturer's source data; no OH/PD value was invented.
