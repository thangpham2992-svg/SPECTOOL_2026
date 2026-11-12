# SpecWebTool v5.8.25 – Material Finish & Preview Update

Base: v5.8.24. Version-specific runtime Material Database override for MRL V02.

- MRL Painted car finishes: Y033D, Y002B, Y116B, Y014B, Y055B (five only).
- MRL Painted ceiling: Y135D, Y136D, Y055B (three only).
- MRL Laminated steel sheet (replaces PPS in MRL only): CP23, CP101, CP141, LA03, LA02 for car wall, door and transom.
- MRL finish/color/pattern lookup tables and source-engine literal expressions are normalized at runtime, including user-imported older MRL spreadsheets.
- The MR database is not changed; painted swatch preview is displayed only for available MR color choices.
- Preview thumbnails come from the three provided catalog screenshots. For Y116B and Y014B (not present in the MRL reference image), the similarly numbered MR colors are shown as reference swatches and labeled accordingly.
- Original bundled XLSX workbooks are retained unmodified as traceable source files; the web application applies its v5.8.25 material overrides at runtime. Do not use the original MRL XLSX alone as evidence of the updated list.

QA: node --check app.js/material-finish.js passed; runtime MRL field/table/preview assertions passed. Full browser Check Spec and exported Excel/PDF were not end-to-end verified in this environment.
