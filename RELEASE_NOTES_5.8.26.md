# v5.8.26 – Car design single-row layout (MR and MRL Version02)

- Group each device on one line: part name, finish, active painted color / laminate pattern.
- Preserve existing SpecForm cells, Excel-linked inputs and listeners.
- Align finish preview immediately to the right of corresponding dropdown.
- When Painted is selected, show the paint color selector; when Laminated is selected, show pattern selector.
- MRL V02 material lists remain as in v5.8.25.

## Pending verification
- The newly supplied YZ-55SN MR cabin-code whitelist cannot be encoded safely without readable code list from the screenshot; the prior mirror constraint remains in force.
- Browser smoke test could not run: local HTTP is blocked by runtime administrator. Static JavaScript syntax checks passed.
