import * as XLSX from "xlsx";

export function cellIsInSqref(cellAddress, sqref) {
  const cell = XLSX.utils.decode_cell(cellAddress);

  return String(sqref ?? "")
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .some((part) => {
      try {
        const range = XLSX.utils.decode_range(part.replace(/\$/g, ""));
        return (
          cell.c >= range.s.c &&
          cell.c <= range.e.c &&
          cell.r >= range.s.r &&
          cell.r <= range.e.r
        );
      } catch {
        return false;
      }
    });
}

export function extendSqrefThroughRow(sqref, firstDataRow, lastRow) {
  const firstDataRowIndex = firstDataRow - 1;
  const lastRowIndex = lastRow - 1;

  return String(sqref ?? "")
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .map((part) => {
      try {
        const range = XLSX.utils.decode_range(part.replace(/\$/g, ""));

        if (range.s.r <= firstDataRowIndex && range.e.r >= firstDataRowIndex) {
          range.e.r = Math.max(range.e.r, lastRowIndex);
        }

        return XLSX.utils.encode_range(range);
      } catch {
        return part;
      }
    })
    .join(" ");
}
