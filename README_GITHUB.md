# Webtool v5.9.3 – iOS Excel workbook & Excel rows 29–30

- Fix workbook relationship lookup for SpecFormExport using raw OOXML instead of iOS DOMParser attribute namespace resolution.
- Preserve the existing PDF export behavior.
- Increase worksheet row heights 29 and 30 in the generated XLSX for both models to avoid text overlap, including the original MR row 29 at 11.4 pt.
- The original source databases remain unchanged; row height correction happens when generating the final .xlsx.

Deploy all files and directories to the GitHub Pages publish folder, not index.html alone. Confirm the updated script query string `?v=5.9.3` in the deployed page.
