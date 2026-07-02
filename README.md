# Version 12 — case-insensitive headers and Additional Pallet rows

Replace these files in the React project:

- `src/config/transformConfig.js`
- `src/models/InputItem.js`
- `src/services/InputWorkbookReader.js`
- `src/App.jsx`

## Changes

- All input-header matching is **case-insensitive but otherwise exact**. For example, `CustomerPO`, `customerpo`, and `CUSTOMERPO` are accepted. `Customer PO` is still a different header because spacing/punctuation are not treated as equivalent.
- The `container` input header is optional and is read case-insensitively.
- In **FTL**, a row whose `container` value is `additional pallet` in any capitalization is retained even when `Old Item code` and `QTY` are blank.
- On an Additional Pallet row, blank `Old Item code`, `QTY`, `L"`, `W"`, `H"`, or `G.W. (LBS)` values become `Error: Additional Pallet` in non-dropdown output fields.
- Freight Class stays blank when it cannot be calculated, because it is a dropdown field.
- Normal rows retain the previous behavior: missing required source values log errors and produce `ERROR` as appropriate.
