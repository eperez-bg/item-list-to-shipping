/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: TemplateLayoutService.js                                             ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Finds the correct template worksheet and header row, then maps each        ║
 * ║ configured field to a template column.                                     ║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ Template headers can shift, so output code should discover fields by       ║
 * ║ their labels instead of hardcoding column letters.                         ║
 * ║                                                                            ║
 * ║ MAIN CLASSES / FUNCTIONS                                                   ║
 * ║ • TemplateLayoutService — Template headers can shift, so output code       ║
 * ║ should discover fields by their labels instead of hardcoding column        ║
 * ║ letters.                                                                   ║
 * ║ • findLayout — The rest of the pipeline needs one verified layout before   ║
 * ║ any XML is edited.                                                         ║
 * ║ • findBestHeaderRow — Headers might not be on row one, so the service      ║
 * ║ needs to evaluate every row objectively.                                   ║
 * ║ • findLastExistingDataRow — When rewriting a template, stale values from   ║
 * ║ older populated rows need to be cleared too.                               ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import * as XLSX from "xlsx";
import { TemplateLayout } from "../models/TemplateLayout";
import { UserFacingError } from "../utils/errors";
import { isBlank, normalizeHeader } from "../utils/text";

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ CLASS: TemplateLayoutService                                               │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Template configuration and required field definitions.
 *
 * Returns
 *   • A reusable layout discovery service.
 *
 * Why / How
 *   Template headers can shift, so output code should discover fields by
 *   their labels instead of hardcoding column letters.
 *   It scores every worksheet/header row against configured fields,
 *   validates all fields exist, and returns a TemplateLayout.
 */
export class TemplateLayoutService {
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ CONSTRUCTOR: constructor                                                   │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • templateConfig and fields.
   *
   * Returns
   *   • A configured TemplateLayoutService.
   *
   * Why / How
   *   Layout discovery must know the requested output sheet preference and the
   *   logical fields expected by the active mode.
   *   It stores the template config and field definitions.
   */
  constructor(templateConfig, fields) {
    this.templateConfig = templateConfig;
    this.fields = fields;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: findLayout                                                         │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Parsed template workbook.
   *
   * Returns
   *   • Verified TemplateLayout or UserFacingError.
   *
   * Why / How
   *   The rest of the pipeline needs one verified layout before any XML is
   *   edited.
   *   It scans eligible sheets, picks the best header match, verifies missing
   *   fields, and constructs TemplateLayout.
   */
  findLayout(workbook) {
    const candidateSheetNames = this.templateConfig.outputSheetName
      ? [this.templateConfig.outputSheetName]
      : workbook.SheetNames;

    let bestCandidate = null;

    for (const sheetName of candidateSheetNames) {
      const worksheet = workbook.Sheets[sheetName];
      if (!worksheet) continue;

      const candidate = this.findBestHeaderRow(worksheet, sheetName);
      if (!bestCandidate || candidate.matchCount > bestCandidate.matchCount) {
        bestCandidate = candidate;
      }
    }

    if (!bestCandidate) {
      throw new UserFacingError("Could not find a usable worksheet in the template workbook.");
    }

    const missingFields = this.fields.filter(
      (field) => bestCandidate.columnsByKey[field.key] === undefined,
    );

    if (missingFields.length > 0) {
      const missing = missingFields
        .map((field) => `${field.header}${field.occurrence ? ` (occurrence ${field.occurrence})` : ""}`)
        .join(", ");

      throw new UserFacingError(
        `Could not find all required yellow-template headers on worksheet "${bestCandidate.sheetName}". Missing: ${missing}.`,
      );
    }

    return new TemplateLayout({
      sheetName: bestCandidate.sheetName,
      headerRow: bestCandidate.headerRow,
      firstDataRow: bestCandidate.headerRow + 1,
      columnsByKey: bestCandidate.columnsByKey,
      lastExistingDataRow: bestCandidate.lastExistingDataRow,
    });
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: findBestHeaderRow                                                  │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Worksheet and sheet name.
   *
   * Returns
   *   • Best candidate layout metadata for that worksheet.
   *
   * Why / How
   *   Headers might not be on row one, so the service needs to evaluate every
   *   row objectively.
   *   It builds normalized header positions per row, maps configured
   *   occurrences, and keeps the row with the most matches.
   */
  findBestHeaderRow(worksheet, sheetName) {
    const rows = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: "",
      raw: false,
      blankrows: true,
    });

    let best = {
      sheetName,
      headerRow: 1,
      matchCount: -1,
      columnsByKey: {},
      lastExistingDataRow: 1,
    };

    rows.forEach((row, rowIndex) => {
      const positionsByHeader = new Map();

      row.forEach((value, columnIndex) => {
        const normalized = normalizeHeader(value);
        if (!normalized) return;

        if (!positionsByHeader.has(normalized)) {
          positionsByHeader.set(normalized, []);
        }
        positionsByHeader.get(normalized).push(columnIndex);
      });

      const columnsByKey = {};
      let matchCount = 0;

      this.fields.forEach((field) => {
        const positions = positionsByHeader.get(normalizeHeader(field.header)) ?? [];
        const occurrence = field.occurrence ?? 1;
        const columnIndex = positions[occurrence - 1];

        if (columnIndex !== undefined) {
          columnsByKey[field.key] = columnIndex;
          matchCount += 1;
        }
      });

      if (matchCount > best.matchCount) {
        best = {
          sheetName,
          headerRow: rowIndex + 1,
          matchCount,
          columnsByKey,
          lastExistingDataRow: this.findLastExistingDataRow(rows, rowIndex + 1, columnsByKey),
        };
      }
    });

    return best;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: findLastExistingDataRow                                            │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Sheet rows, 1-based header row, mapped columns.
   *
   * Returns
   *   • 1-based last used data row.
   *
   * Why / How
   *   When rewriting a template, stale values from older populated rows need
   *   to be cleared too.
   *   It scans rows below headers and remembers the last row containing any
   *   mapped field value.
   */
  findLastExistingDataRow(rows, headerRowIndex, columnsByKey) {
    let lastRow = headerRowIndex;
    const mappedColumns = Object.values(columnsByKey);

    for (let rowIndex = headerRowIndex; rowIndex < rows.length; rowIndex += 1) {
      const hasMappedValue = mappedColumns.some(
        (columnIndex) => !isBlank(rows[rowIndex]?.[columnIndex]),
      );

      if (hasMappedValue) lastRow = rowIndex + 1;
    }

    return lastRow;
  }
}
