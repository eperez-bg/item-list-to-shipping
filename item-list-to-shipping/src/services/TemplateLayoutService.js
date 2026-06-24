import * as XLSX from "xlsx";
import { TemplateLayout } from "../models/TemplateLayout";
import { UserFacingError } from "../utils/errors";
import { isBlank, normalizeHeader } from "../utils/text";

export class TemplateLayoutService {
  constructor(templateConfig, fields) {
    this.templateConfig = templateConfig;
    this.fields = fields;
  }

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
