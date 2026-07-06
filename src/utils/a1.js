/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: a1.js                                                                ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Contains small helpers for working with Excel A1-style addresses and       ║
 * ║ data-validation ranges.                                                    ║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ Data validation can apply to several space-separated ranges, so callers    ║
 * ║ need one safe membership test.                                             ║
 * ║                                                                            ║
 * ║ MAIN CLASSES / FUNCTIONS                                                   ║
 * ║ • cellIsInSqref — Data validation can apply to several space-separated     ║
 * ║ ranges, so callers need one safe membership test.                          ║
 * ║ • extendSqrefThroughRow — Template dropdowns must extend to generated      ║
 * ║ output rows without changing unrelated validation ranges.                  ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import * as XLSX from "xlsx";

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: cellIsInSqref                                                    │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • A1 cell address and sqref string.
 *
 * Returns
 *   • Boolean indicating whether the cell is covered.
 *
 * Why / How
 *   Data validation can apply to several space-separated ranges, so callers
 *   need one safe membership test.
 *   It decodes the cell and each sqref part, returning true for any
 *   containing range while ignoring invalid parts.
 */
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

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: extendSqrefThroughRow                                            │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Original sqref, 1-based first data row, 1-based last row.
 *
 * Returns
 *   • Updated space-separated sqref string.
 *
 * Why / How
 *   Template dropdowns must extend to generated output rows without changing
 *   unrelated validation ranges.
 *   It decodes each sqref range and expands only ranges that contain the
 *   first data row.
 */
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
