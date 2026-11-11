# v5.9.2 – iOS XLSX relationship fix

Fixes `No workbook relationship for Export sheet` in the XLSX generator by finding relationships through XML namespace-aware traversal rather than `.children`, which may behave inconsistently in WebKit XML documents. Also handles UTF-8 BOM and a missing qualified r:id attribute.

Deploy the *entire* folder contents to the GitHub Pages publication folder, then verify v5.9.2 in the header. PDF export, database, and engineering validation were not changed.

Test on a real iOS device is still required; JS syntax and XLSX reference IDs were checked locally.
