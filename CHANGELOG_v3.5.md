# v3.5 — MRL Overhead / Service floor

- **MRL OH:** Use original `SpecForm!F37` formula against `DataTable!BG8:BS43` for live minimum OH guidance. Raise OH to the minimum only if the current input falls below the requirement when capacity, speed, travel, or regulation changes. Do not overwrite larger manual OH. Show unsupported combinations explicitly.
- **Service floor:** Free text for both MR and MRL; entirely remove the legacy C# validation of floor names, floor-count equality with Stops, and associated derived fixture checks. `Stops and opens` remains independently checked.
- No numerical OH source values were added or guessed. Both bundled workbook databases are retained as supplied in v3.4.
