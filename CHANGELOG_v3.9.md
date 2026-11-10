# v3.9 — Precise OH/PD/AH/BH errors + blue exported headings

- Both Excel and A4 PDF exports retain pale blue fill on Roman-numbered specification sections (I, II, III, IV); regular Feature rows remain free of brown/tan highlighting (changes from v3.8).
- MR and MRL Check Spec: error messages for actual OH/PD/AH/BH values compared with workbook-calculated minima or maxima. Display actual value and required limit with mm and Vietnamese number formatting.
- MRL DataTable D522 now reports AH and BH independently and navigates to the correct cell (formerly aggregated to AH).
- MRL DataTable D523 reports OH and PD independently; absent source lookup is explicitly CHECK LAYOUT and never supplies an invented limit.
- MR DataTable AP24/AP25 dynamically use effective OH (B529/E521) and PD (E524/E523) thresholds; C# Step1 also reports out-of-range AH/BH and pit max.
- Check Spec panel expands only dimensional errors with the exact limit, leaving other messages short. Duplicate high-level aggregate B209 issues are suppressed when precise MRL dimensions are available.
- Unmodified Excel Database, formulas, other technical conditions, and export field values.
- MRL upper AH/BH range: use actual TblK5_Capacity AH-Max/BH-Max columns (AH-Max-FIRE/BH-Max-FIRE when fire doors are selected). No invented limits; missing values stay unverified.
