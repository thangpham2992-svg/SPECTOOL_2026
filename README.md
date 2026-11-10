# SpecWebTool v4.7 — Fireproof door options and MR other-floor export

## Fireproof door — both models
- Not supplied (Không cung cấp)
- GHCL E30
- EN81-58(II)

GHCL E30 uses the **Not supplied** AH/BH lookup branch for both MR and MRL. EN81-58(II) continues to use the fire-rated AH/BH lookup where source data requires it. MRL retains the version 4.6 156-row four-counterweight database including corrected AH/BH originals.

## Export corrections
- NEXIEZ-MR other-floor PIV and HBV now include the same Vietnamese/English material/case description as the main floor, with each floor's own model code.
- The fireproof door section exports the selected option, including a distinct GHCL E30 description. **Not supplied** is not automatically shown as E30.
- Source Excel workbooks and embedded offline workbooks are updated.

For offline usage, unzip and open `index.html`. The `database` directory contains both XLSX databases.
