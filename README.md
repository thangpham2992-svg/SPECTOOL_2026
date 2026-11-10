# SpecWebTool v5.0 — ITP 24V 5 lines + inline Features

- Added **24V – 5 Lines – 2 Master Stations** and **24V – 5 Lines – 3 Master Stations** to **NEXIEZ-MR** / 2C-2BC. Both options were already present in NEXIEZ-MRL Version02 and are preserved.
- The new choices are injected into the live workbook DataTable source for existing bundled databases and user-imported MR databases. The source Excel files in `database/` retain their original bytes; new MR choices apply inside this Webtool and Export.
- ITP is now displayed only under **Features → Tín hiệu / Signal**, immediately beneath the ITP checkbox (no pinned panel). **ITV Cable** remains in the Signal Feature list.
- The first valid ITP choice is still saved automatically on page load, even before opening the Signal tab. Manual choices are preserved where valid; changing operation or FE triggers revalidation.
- Check Spec and Export use the chosen ITP type (including both new 24V variants).
- All other MR/MRL model rules, photos and reports are retained from v4.9.

Unzip and open `index.html` for offline usage.
