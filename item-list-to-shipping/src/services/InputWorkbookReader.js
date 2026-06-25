import * as XLSX from "xlsx";
import { InputItem } from "../models/InputItem";
import { InputShipment } from "../models/InputShipment";
import { InputSkid } from "../models/InputSkid";
import { UserFacingError } from "../utils/errors";
import { ERROR_VALUE, ImportIssueReporter } from "../utils/importIssues";
import { getCellDisplayValue, isBlank, toFiniteNumber } from "../utils/text";

export class InputWorkbookReader {
  constructor(inputConfig) {
    this.config = inputConfig;
  }

  async readFile(file) {
    return this.readArrayBuffer(await file.arrayBuffer());
  }

  readArrayBuffer(arrayBuffer) {
    const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
    const sheetName = workbook.SheetNames[this.config.sheetIndex];

    if (!sheetName) {
      this.throwCriticalInputError({
        type: "INPUT_SHEET_NOT_FOUND",
        location: "Workbook",
        field: "Input worksheet",
        message: "The uploaded workbook does not have the expected first worksheet.",
        resolution: "Upload a workbook with the input sheet as its first worksheet.",
      });
    }

    const worksheet = workbook.Sheets[sheetName];
    const usedRange = worksheet?.["!ref"];

    if (!usedRange) {
      this.throwCriticalInputError({
        type: "EMPTY_INPUT_WORKSHEET",
        location: sheetName,
        field: "Input worksheet",
        message: "The uploaded worksheet is empty.",
        resolution: "Upload a worksheet containing the required input headers and item rows.",
      });
    }

    const issues = [];
    const reporter = new ImportIssueReporter(issues);
    const merges = worksheet["!merges"] ?? [];
    const range = XLSX.utils.decode_range(usedRange);
    const headerMap = this.findRequiredHeaderColumns(worksheet, range, sheetName);
    const groupedItems = new Map();
    let hasReachedData = false;

    for (
      let rowIndex = headerMap.headerRowIndex + 1;
      rowIndex <= range.e.r;
      rowIndex += 1
    ) {
      const rowNumber = rowIndex + 1;

      const rawOldItemCode = this.readEffectiveValue(
        worksheet,
        headerMap.columns.oldItemCode,
        rowNumber,
        merges,
      );
      const rawQuantity = this.readEffectiveValue(
        worksheet,
        headerMap.columns.quantity,
        rowNumber,
        merges,
      );

      // Ignore completely blank rows. Before actual item data begins, also
      // ignore labels/notes that do not have a quantity.
      if (isBlank(rawOldItemCode) && isBlank(rawQuantity)) continue;
      if (!hasReachedData && isBlank(rawQuantity)) continue;

      hasReachedData = true;

      const customerPo = this.valueOrError({
        value: this.readEffectiveValue(
          worksheet,
          headerMap.columns.customerPo,
          rowNumber,
          merges,
        ),
        columnIndex: headerMap.columns.customerPo,
        rowNumber,
        field: "Customer PO",
        reporter,
      });

      const oldItemCode = this.valueOrError({
        value: rawOldItemCode,
        columnIndex: headerMap.columns.oldItemCode,
        rowNumber,
        field: "Old Item Code",
        reporter,
      });

      const quantity = this.valueOrError({
        value: rawQuantity,
        columnIndex: headerMap.columns.quantity,
        rowNumber,
        field: "Quantity",
        reporter,
      });

      const targetStore = this.readEffectiveValue(
        worksheet,
        headerMap.columns.targetStore,
        rowNumber,
        merges,
      );

      if (isBlank(targetStore)) {
        const location = this.addressFor(
          headerMap.columns.targetStore,
          rowNumber,
        );

        reporter.record({
          type: "MISSING_INPUT_VALUE",
          location,
          field: "Target Store",
          message: `Input cell ${location} is blank.`,
          resolution:
            'Destination is a template dropdown, so it will be left blank. The second Target PO# field will receive "ERROR".',
        });
      }

      const pickupNumber = this.valueOrError({
        value: this.readEffectiveValue(
          worksheet,
          headerMap.columns.pickupNumber,
          rowNumber,
          merges,
        ),
        columnIndex: headerMap.columns.pickupNumber,
        rowNumber,
        field: "Pickup Number",
        reporter,
      });

      const skidRange = this.findSkidRangeForRow(
        rowIndex,
        merges,
        headerMap.skidGroupingColumnIndexes,
      );
      const skidKey = `${skidRange.startRow}:${skidRange.endRow}`;

      const item = new InputItem({
        sourceRow: rowNumber,
        customerPo,
        targetStore,
        oldItemCode,
        quantity,
        pickupNumber,
      });

      if (!groupedItems.has(skidKey)) {
        groupedItems.set(skidKey, { range: skidRange, items: [] });
      }

      groupedItems.get(skidKey).items.push(item);
    }

    if (groupedItems.size === 0) {
      this.throwCriticalInputError({
        type: "NO_INPUT_ITEMS_FOUND",
        location: `${sheetName}!${headerMap.headerRowNumber}:${headerMap.headerRowNumber}`,
        field: "Input item rows",
        message:
          "No item rows were found below the required input header row.",
        resolution:
          "Check that item rows contain values in both Old Item code and QTY.",
      });
    }

    const skids = Array.from(groupedItems.values())
      .sort((left, right) => left.range.startRow - right.range.startRow)
      .map(({ range: skidRange, items }) =>
        this.createSkid(
          worksheet,
          merges,
          skidRange,
          items,
          reporter,
          headerMap.columns,
        ),
      );

    return new InputShipment({ skids, issues });
  }

  /*
    Finds the one row that contains every required header. Header matching is:
      - exact
      - case-sensitive
      - no partial/similar-header matching

    The resulting column indexes are used everywhere else, so deleted or moved
    input columns do not affect importing.
  */
  findRequiredHeaderColumns(worksheet, range, sheetName) {
    const configuredHeaders = this.config.headers ?? {};
    const requiredFields = Object.keys(configuredHeaders);
    const requiredHeaderTexts = Object.values(configuredHeaders);

    if (requiredFields.length === 0) {
      this.throwCriticalInputError({
        type: "INPUT_HEADER_CONFIGURATION_ERROR",
        location: "Application configuration",
        field: "Input headers",
        message: "No required input headers are configured.",
        resolution: "Add the required exact header names in transformConfig.js.",
      });
    }

    const duplicateConfiguredHeaders = requiredHeaderTexts.filter(
      (header, index) => requiredHeaderTexts.indexOf(header) !== index,
    );

    if (duplicateConfiguredHeaders.length > 0) {
      this.throwCriticalInputError({
        type: "DUPLICATE_INPUT_HEADER_CONFIGURATION",
        location: "Application configuration",
        field: "Input headers",
        message:
          `Each configured input header must be unique. Duplicate values: ` +
          `${[...new Set(duplicateConfiguredHeaders)].join(", ")}.`,
        resolution: "Correct the headers object in transformConfig.js.",
      });
    }

    const matchingColumnsByRow = new Map();

    for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
      const matchesForRow = new Map();

      for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
        const cell = worksheet[
          XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })
        ];
        const cellText = getCellDisplayValue(cell).trim();

        if (!requiredHeaderTexts.includes(cellText)) {
          continue;
        }

        if (!matchesForRow.has(cellText)) {
          matchesForRow.set(cellText, []);
        }

        matchesForRow.get(cellText).push(columnIndex);
      }

      if (matchesForRow.size > 0) {
        matchingColumnsByRow.set(rowIndex, matchesForRow);
      }
    }

    const candidateHeaderRows = Array.from(matchingColumnsByRow.entries())
      .filter(([, matchesForRow]) =>
        requiredHeaderTexts.every(
          (headerText) => (matchesForRow.get(headerText) ?? []).length === 1,
        ),
      )
      .map(([rowIndex, matchesForRow]) => ({ rowIndex, matchesForRow }));

    if (candidateHeaderRows.length === 0) {
      const foundHeaderLocations = this.describeFoundHeaderLocations(
        matchingColumnsByRow,
      );

      this.throwCriticalInputError({
        type: "REQUIRED_INPUT_HEADERS_NOT_FOUND",
        location: sheetName,
        field: "Input headers",
        message:
          "Could not find one row containing every required exact input header. " +
          `Required headers: ${requiredHeaderTexts.join(", ")}. ` +
          `Found exact header cells: ${foundHeaderLocations || "none"}.`,
        resolution:
          "Confirm the input workbook contains all required header names exactly as configured. No similar-header matching is used.",
      });
    }

    if (candidateHeaderRows.length > 1) {
      const rowNumbers = candidateHeaderRows
        .map(({ rowIndex }) => rowIndex + 1)
        .join(", ");

      this.throwCriticalInputError({
        type: "AMBIGUOUS_INPUT_HEADER_ROW",
        location: sheetName,
        field: "Input headers",
        message:
          `More than one row contains all required exact input headers: rows ${rowNumbers}.`,
        resolution:
          "Keep only one complete input-header row in the uploaded worksheet.",
      });
    }

    const [{ rowIndex: headerRowIndex, matchesForRow }] = candidateHeaderRows;
    const columns = {};

    requiredFields.forEach((field) => {
      const headerText = configuredHeaders[field];
      columns[field] = matchesForRow.get(headerText)[0];
    });

    const skidGroupingColumnIndexes = (this.config.skidGroupingFields ?? [])
      .map((field) => columns[field])
      .filter((columnIndex) => Number.isInteger(columnIndex));

    if (skidGroupingColumnIndexes.length === 0) {
      this.throwCriticalInputError({
        type: "SKID_GROUPING_CONFIGURATION_ERROR",
        location: "Application configuration",
        field: "Skid grouping headers",
        message:
          "None of the configured skidGroupingFields resolved to a required input header.",
        resolution:
          "Use field keys from input.headers, such as length, width, and height.",
      });
    }

    console.info("[Fuse Order Template Filler]", {
      type: "INPUT_HEADERS_RESOLVED",
      worksheet: sheetName,
      headerRow: headerRowIndex + 1,
      headers: Object.fromEntries(
        Object.entries(columns).map(([field, columnIndex]) => [
          field,
          {
            header: configuredHeaders[field],
            column: XLSX.utils.encode_col(columnIndex),
          },
        ]),
      ),
    });

    return {
      headerRowIndex,
      headerRowNumber: headerRowIndex + 1,
      columns,
      skidGroupingColumnIndexes,
    };
  }

  describeFoundHeaderLocations(matchingColumnsByRow) {
    const locations = [];

    matchingColumnsByRow.forEach((matchesForRow, rowIndex) => {
      matchesForRow.forEach((columnIndexes, headerText) => {
        columnIndexes.forEach((columnIndex) => {
          locations.push(
            `${headerText} at ${XLSX.utils.encode_col(columnIndex)}${rowIndex + 1}`,
          );
        });
      });
    });

    return locations.join("; ");
  }

  createSkid(worksheet, merges, skidRange, items, reporter, columns) {
    const sourceRow = skidRange.startRow;

    return new InputSkid({
      sourceStartRow: skidRange.startRow,
      sourceEndRow: skidRange.endRow,
      length: this.readPositiveNumberOrError({
        worksheet,
        merges,
        columnIndex: columns.length,
        rowNumber: sourceRow,
        field: "Length",
        reporter,
      }),
      width: this.readPositiveNumberOrError({
        worksheet,
        merges,
        columnIndex: columns.width,
        rowNumber: sourceRow,
        field: "Width",
        reporter,
      }),
      height: this.readPositiveNumberOrError({
        worksheet,
        merges,
        columnIndex: columns.height,
        rowNumber: sourceRow,
        field: "Height",
        reporter,
      }),
      grossWeight: this.readPositiveNumberOrError({
        worksheet,
        merges,
        columnIndex: columns.grossWeight,
        rowNumber: sourceRow,
        field: "Gross Weight",
        reporter,
      }),
      items,
    });
  }

  readPositiveNumberOrError({
    worksheet,
    merges,
    columnIndex,
    rowNumber,
    field,
    reporter,
  }) {
    const rawValue = this.readEffectiveValue(
      worksheet,
      columnIndex,
      rowNumber,
      merges,
    );
    const numericValue = toFiniteNumber(rawValue);
    const location = this.addressFor(columnIndex, rowNumber);

    if (numericValue !== null && numericValue > 0) {
      return numericValue;
    }

    const missing = isBlank(rawValue);

    reporter.record({
      type: missing ? "MISSING_INPUT_VALUE" : "INVALID_NUMERIC_INPUT",
      location,
      field,
      message: missing
        ? `Input cell ${location} is blank.`
        : `Input cell ${location} must be a positive number, but contains "${displayText(rawValue)}".`,
      resolution: `The non-dropdown ${field} output cell will receive "ERROR". Freight Class is a dropdown and will be left blank.`,
    });

    return ERROR_VALUE;
  }

  valueOrError({ value, columnIndex, rowNumber, field, reporter }) {
    if (!isBlank(value)) {
      return value;
    }

    const location = this.addressFor(columnIndex, rowNumber);

    reporter.record({
      type: "MISSING_INPUT_VALUE",
      location,
      field,
      message: `Input cell ${location} is blank.`,
      resolution: `The non-dropdown ${field} output cell will receive "ERROR".`,
    });

    return ERROR_VALUE;
  }

  findSkidRangeForRow(rowIndex, merges, groupingColumnIndexes) {
    const matchingRanges = merges
      .filter((merge) => {
        const vertical = merge.e.r > merge.s.r;
        const containsRow = merge.s.r <= rowIndex && rowIndex <= merge.e.r;
        const touchesGroupingColumn = groupingColumnIndexes.some(
          (columnIndex) => merge.s.c <= columnIndex && columnIndex <= merge.e.c,
        );

        return vertical && containsRow && touchesGroupingColumn;
      })
      .sort((left, right) => {
        const leftLength = left.e.r - left.s.r;
        const rightLength = right.e.r - right.s.r;
        return rightLength - leftLength;
      });

    const bestRange = matchingRanges[0];

    if (!bestRange) {
      return { startRow: rowIndex + 1, endRow: rowIndex + 1 };
    }

    return { startRow: bestRange.s.r + 1, endRow: bestRange.e.r + 1 };
  }

  readEffectiveValue(worksheet, columnIndex, rowNumber, merges) {
    const cell = this.readEffectiveCell(
      worksheet,
      columnIndex,
      rowNumber,
      merges,
    );

    // .v preserves raw values such as Z. .w is only a formatted-text fallback.
    return cell?.v ?? cell?.w ?? "";
  }

  readEffectiveCell(worksheet, columnIndex, rowNumber, merges) {
    const address = this.addressFor(columnIndex, rowNumber);
    const directCell = worksheet[address];

    if (!isBlank(getCellDisplayValue(directCell))) {
      return directCell;
    }

    const coordinate = XLSX.utils.decode_cell(address);
    const merge = merges.find(
      (candidate) =>
        candidate.s.c <= coordinate.c &&
        coordinate.c <= candidate.e.c &&
        candidate.s.r <= coordinate.r &&
        coordinate.r <= candidate.e.r,
    );

    if (!merge) {
      return directCell;
    }

    return worksheet[XLSX.utils.encode_cell(merge.s)] ?? directCell;
  }

  addressFor(columnIndex, rowNumber) {
    return `${XLSX.utils.encode_col(columnIndex)}${rowNumber}`;
  }

  throwCriticalInputError({ type, location, field, message, resolution }) {
    console.error("[Fuse Order Template Filler]", {
      type,
      location,
      field,
      message,
      resolution,
    });

    throw new UserFacingError(message);
  }
}

function displayText(value) {
  return String(value ?? "").trim();
}
