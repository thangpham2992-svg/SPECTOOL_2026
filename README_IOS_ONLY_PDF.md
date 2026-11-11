# Webtool v5.8.5 — PC/Android format restoration

Base: v5.8.3, NOT v5.8.4.

- PC/Android: print-layout.js uses exactly the prior v5.8.3 HTML/CSS; no global borders, paper zoom, page margins, or layout changes from v5.8.4.
- iOS only: native borders (if and only if present in original SpecFormExport style) are increased from 0.45px to 0.8px. Print color adjustment enabled.
- XLSX export logic, workbook database, technical specs, and iOS XML parsing fix all unchanged from v5.8.3.
- PDF/Paper URL, date and time may be inserted by the iOS print UI itself and cannot be guaranteed hidden by CSS.
- Test iPhone/iPad plus Android/PC before deploying into production.
