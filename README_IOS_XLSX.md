# v5.9.0 — iOS Excel export compatibility

- On iPhone/iPad, Excel Export opens a download tab synchronously during the user click, then populates its XLSX link once the workbook has been generated. Users tap **Tải Excel (.xlsx)** and use Share > Save to Files if the file is previewed.
- If pop-ups are blocked, an in-page XLSX link is displayed.
- The temporary Blob URL is retained for 5 minutes rather than immediately revoked.
- Errors in Excel generation are now shown with detail rather than generic Check Spec wording.
- PDF flow, underlying source workbooks, technical rules and feature conditions have not been intentionally changed.
- Upload ALL contents of this directory to the GitHub Pages publishing root. Check that website header reads v5.9.0.
- Actual iOS Safari/Chrome/Edge device testing is required before considering the fix confirmed.
