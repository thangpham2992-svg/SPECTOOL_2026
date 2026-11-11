# SpecWebTool v5.7 — NEXIEZ-MRL Version02 Hall Lanterns and feature interlocks

## Usage
Unzip all files to one folder, then open `index.html` in a browser. Keep `assets/hall-lanterns/` beside the other application files.

## NEXIEZ-MRL Version02
- In IV. Entrance design / Hall indicators & buttons, the three floor cards each have one Hall Lanterns dropdown immediately below PIH/PID: Top floor, Other floors, Main floor.
- 14 Hall Lantern models from Sales Manual S-1-U4010-E, Table 5.6.2 (p.27), with original product image thumbnails and a name/image gallery: HLV-A16S, HLV-E65, HLV-E66, HLH-A16S, HLV1-A80S, HLV3-A80S, HLV1-A90S, HLV3-A90S, HLV1-E62, HLV3-E62, HLH1-A80S, HLH3-A80S, HLH1-A90S and HLH3-A90S.
- When one or more Hall Lanterns are selected, FHL is automatically selected. Without Hall Lanterns, FHL is not allowed.
- AECC and AECH are mutually exclusive; AECH and AIL are mutually exclusive.
- AIL requires Hall Lantern and SAI-2200C group control, following Manual Table 6.4.2.
- Check Spec points to the actual conflicting Feature. Export includes each selected Hall Lantern code in the corresponding floor's entrance-design text (bilingual), retaining the existing SpecFormExport layout.

## Excel source and model isolation
- `database/SpecToolGQXL2P.xlsx` is the updated MRL Version02 source. It adds `HallLantern_V02` with catalog and Excel-check formulas for the user-editable source selections. Existing technical formulas and other sheets remain intact.
- `database/SpecToolGXLM2P.xlsx` is the unchanged MR v5.5 source. MR does not use the MRL lanterns or its feature conditions.
- The embedded XLSX and calculation dataset were synchronized so this package works without a server.

## QA
Verified JS syntax, 3 MRL selectors (15 choices each), 14 gallery items, FHL auto selection, MR unaffected, feature interlocks, and exported model text using a Chromium test harness with inlined scripts (the test environment blocks local file navigation). Full saved-XLSX export and external Excel recalc should be verified with a production workbook.

## v5.7 SpecFormExport — Hall Lantern housing / faceplate
- The chosen Hall Lantern is printed in each applicable floor block (Top / Other / Main), preserving any PIV description, with one Vietnamese and one English material sentence.
- SUS-HL (SUS304) applies to A16S, E65, E66, A80S and E62 codes; black plastic applies to A90S codes. Material names follow Manual Table 5.6.2 p.27.
- Selecting a Hall Lantern when PIV is Not supplied no longer prints a misleading standalone “Not supplied” message in its export row.
- XLSX output rows expand to fit bilingual content; PDF is generated from that exact XLSX SpecFormExport worksheet.
- Unselected Hall Lanterns do not appear in Export. Source .xlsx Feature and technical data are unchanged.
