# SpecWebTool v4.9 — ITP defaults to first valid option

## New in v4.9
- ITP dropdown automatically selects the first valid choice on initial load for both MR and MRL.
- When users explicitly select a different ITP choice, it is preserved across UI re-renders while valid.
- Group Control/FE changes select the first eligible option if the earlier one is no longer valid.
- The automatic choice is written into the working spec state so Check Spec/Export sees it. No false ITP Required warning for an untouched valid dropdown.
- Existing v4.8 features/databases remain unchanged.

## Changes
- JJ (mm) standard first, then optional configurations marked `(Option)` in dropdown.
- JJ cell values remain numeric for formulas, shaft AH/BH lookup, Excel Export.
- Eight optional JJ combinations from Version02 car-size table, including P11-CO 825 kg with 800 (Option) and P11-2S with 1100 (Option).
- MRL COP failures report COP fields: D59 button, D60 location, D61 code, D62 wheelchair COP.
- Aggregated failure conditions (D512 / D527 / B209 / B210 / B206) no longer create a spurious Capacity error.
- COP button, location and code validated using original workbook rules.
- All v4.7 changes including GHCL E30, counterweight AH/BH and MR other-floor Export retained.

Use offline: extract the ZIP and open `index.html`. Included databases are in `database/`.
