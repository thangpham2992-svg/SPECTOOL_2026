# SpecWebTool v5.8.55 — Optional technical details export

- New checkbox before Check Spec / Excel / PDF buttons: “Trích xuất chi tiết động cơ / cáp kéo”.
- Default unchecked: no optional Motor/Pulley/Rope rows inserted for either model; Main Code stays immediately after Control system.
- Checked: MR inserts Motor, Pulley, Hoisting rope; MRL V02 inserts Motor only, in the established position after Control system.
- Excel and PDF share the same generated XLSX workbook, so the switch applies to both consistently.
- Existing Check Spec and capacity/speed lookup remain unchanged.
- Previous optional gray row highlighting is applied only if rows are inserted.

## GitHub compact release
Merged 13 sequential classic JavaScript sources into spec-core.bundle.js, preserving order. No assets or databases removed.
