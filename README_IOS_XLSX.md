# v5.8.2 – iOS XLSX export compatibility (based strictly on v5.8.1)

- Core SpecFormExport generator, styles, merged cells, hidden D/E, excluded X rows, PDF and technical databases are unchanged from v5.8.1.
- iOS XML namespace fallback for workbook-to-export sheet relationship.
- iOS XLSX presents a Save/Share button after file creation (Web Share API); a normal Blob download link is offered as a fallback.
- No window.open or Data URL for Excel.
- Non-iOS download remains direct; no change to PDF.
- Requires validation on a physical iPhone/iPad; JavaScript and XLSX package compatibility checks alone cannot prove it.
