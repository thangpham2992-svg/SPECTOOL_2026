# v5.9.1 iOS Excel fix

- No new blank window/tab when exporting XLSX on iOS.
- Generate XLSX on the original page; show a dedicated ready-to-save panel.
- Prefer native Web Share API with XLSX File on a separate user tap (Save to Files).
- If the API is unavailable, provide a Data URL download link in the same page.
- iOS can restrict Data URL downloads: native sharing is preferable, and actual iOS testing is required.
- Keep PDF export and non-iOS Excel behavior unchanged.
- Upload all files to GitHub Pages publishing root, then hard refresh.
