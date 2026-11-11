# SPECTOOL 2026 v5.8.6 — Auto floor granite weight

Based directly on v5.8.5; PDF/XLSX export modules are unchanged.

- NEXIEZ-MR flooring weight at SpecForm D60.
- NEXIEZ-MRL Version02 flooring weight at SpecForm D56.
- Automatic on load and changes to flooring type, AA, BB, JJ or capacity.
- Excel-equivalent formula: If Local granite (20 mm), add 5 kg when JJ < 1000 mm, otherwise 10 kg, plus ROUND(AA*BB*0.02*2.7*0.001,0). Other flooring: 0 kg.
- Weight field remains manually editable. User-entered weight is preserved until a dependent technical input changes; 'Tính lại / Recalculate' restores the calculated weight.
- Original Excel source files already have this formula and are included unchanged in database/.
- XLSX/PDF formatting and export scripts are completely unchanged.
