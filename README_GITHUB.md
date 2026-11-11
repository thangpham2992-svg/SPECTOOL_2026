# SpecWebTool v5.8.1 — GitHub Pages deployment

**Main fix:** GitHub Pages may cache a stale `spec-conditional.js`. v5.8 changed the MR conditional-formatting LightUp rules but `index.html` still referenced that script without versioning. This release sets `?v=5.8.1` on every local JS/CSS and adds a safety check in `app.js` to keep the ceiling material unlocked for `L200S`, `L210S`, `N300S`.

MR suffix-S ceiling material: **SUS-HL only**, selectable and not greyed out. Normal MR models retain existing material choices. NEXIEZ-MRL Version02 unchanged.

## Deploy (GitHub Pages)

1. Unzip `SpecWebTool_v5.8.1_GitHub_Pages.zip` on your computer.
2. In the repository, find the branch and folder configured at **Settings > Pages > Build and deployment** (usually `main` + `/(root)` or `/docs`).
3. Replace **all** files/folders from the zip at that Pages publish location — **not just `index.html`**. Preserve `assets/`, `database/`, and `vendor/`.
4. Commit and push. Check the latest Pages deployment in **Actions** and open the website. The top-right header should show **v5.8.1**.
5. Refresh with **Ctrl+Shift+R** or open an incognito tab. If still stale, inspect DevTools > Network and verify `spec-conditional.js?v=5.8.1` and `app.js?v=5.8.1` return status **200**, not **404**.

## Tests

Select **NEXIEZ-MR**, choose ceiling model `N300S`, `L210S`, or `L200S`. The ceiling finish dropdown should show **only SUS-HL**, stay editable, and Check Spec should not report an invalid paint-color error. Switch back to a normal painted ceiling and ensure its paint dropdown is restored.

The archive includes an updated `database/SpecToolGXLM2P.xlsx` synced to source v5.8. The corresponding source workbook also remains in the root for reference.

`version.txt` can be opened under your Pages site to verify deployed content (e.g. `https://user.github.io/repository/version.txt`).
