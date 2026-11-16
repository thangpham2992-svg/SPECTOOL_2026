# v5.8.77 — GitHub Pages / offline consistency

- Prevented automatic reloading of `database/SpecToolGQXL2P.xlsx` and `database/SpecToolGXLM2P.xlsx` at initial page load or model change. Previously these remote workbooks could overwrite the embedded, patched database and produce different SETS-L results on GitHub Pages compared with local `index.html`.
- Kept a manual Reload Database action for intentional external workbook import.
- Retained all SETS-L validation (including speed/travel constraints), as well as changes from v5.8.76.
- Cache-busted script links to v5.8.77.
