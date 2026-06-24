import * as XLSX from "xlsx";
import { InputItem } from "../models/InputItem";
import { InputShipment } from "../models/InputShipment";
import { InputSkid } from "../models/InputSkid";
import { assertUser } from "../utils/errors";
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

    assertUser(
      sheetName,
      "The uploaded workbook does not have the expected first worksheet.",
    );

    const worksheet = workbook.Sheets[sheetName];
    const usedRange = worksheet["!ref"];

    assertUser(usedRange, "The uploaded worksheet is empty.");

    const issues = [];
    const reporter = new ImportIssueReporter(issues);
    const merges = worksheet["!merges"] ?? [];
    const range = XLSX.utils.decode_range(usedRange);
    const groupedItems = new Map();
    let hasReachedData = false;

    for (
      let rowIndex = this.config.firstPossibleDataRow - 1;
      rowIndex <= range.e.r;
      rowIndex += 1
    ) {
      const rowNumber = rowIndex + 1;

      const rawOldItemCode = this.readEffectiveValue(
        worksheet,
        this.config.columns.oldItemCode,
        rowNumber,
        merges,
      );
      const rawQuantity = this.readEffectiveValue(
        worksheet,
        this.config.columns.quantity,
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
          this.config.columns.customerPo,
          rowNumber,
          merges,
        ),
        column: this.config.columns.customerPo,
        rowNumber,
        field: "Customer PO",
        reporter,
      });

      const oldItemCode = this.valueOrError({
        value: rawOldItemCode,
        column: this.config.columns.oldItemCode,
        rowNumber,
        field: "Old Item Code",
        reporter,
      });
      const quantity = this.valueOrError({
        value: rawQuantity,
        column: this.config.columns.quantity,
        rowNumber,
        field: "Quantity",
        reporter,
      });

      const targetStore = this.readEffectiveValue(
        worksheet,
        this.config.columns.targetStore,
        rowNumber,
        merges,
      );

      if (isBlank(targetStore)) {
        reporter.record({
          type: "MISSING_INPUT_VALUE",
          location: `${this.config.columns.targetStore}${rowNumber}`,
          field: "Target Store",
          message: `Input cell ${this.config.columns.targetStore}${rowNumber} is blank.`,
          resolution:
            'Destination is a template dropdown, so it will be left blank. The second Target PO# field will receive "ERROR".',
        });
      }

      const pickupNumber = this.valueOrError({
        value: this.readEffectiveValue(
          worksheet,
          this.config.columns.pickupNumber,
          rowNumber,
          merges,
        ),
        column: this.config.columns.pickupNumber,
        rowNumber,
        field: "Pickup Number",
        reporter,
      });

      const skidRange = this.findSkidRangeForRow(rowIndex, merges);
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

    assertUser(
      groupedItems.size > 0,
      "No item rows were found in the uploaded workbook.",
    );

    const skids = Array.from(groupedItems.values())
      .sort((left, right) => left.range.startRow - right.range.startRow)
      .map(({ range: skidRange, items }) =>
        this.createSkid(worksheet, merges, skidRange, items, reporter),
      );

    return new InputShipment({ skids, issues });
  }

  createSkid(worksheet, merges, skidRange, items, reporter) {
    const sourceRow = skidRange.startRow;

    return new InputSkid({
      sourceStartRow: skidRange.startRow,
      sourceEndRow: skidRange.endRow,
      length: this.readPositiveNumberOrError({
        worksheet,
        merges,
        column: this.config.columns.length,
        rowNumber: sourceRow,
        field: "Length",
        reporter,
      }),
      width: this.readPositiveNumberOrError({
        worksheet,
        merges,
        column: this.config.columns.width,
        rowNumber: sourceRow,
        field: "Width",
        reporter,
      }),
      height: this.readPositiveNumberOrError({
        worksheet,
        merges,
        column: this.config.columns.height,
        rowNumber: sourceRow,
        field: "Height",
        reporter,
      }),
      grossWeight: this.readPositiveNumberOrError({
        worksheet,
        merges,
        column: this.config.columns.grossWeight,
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
    column,
    rowNumber,
    field,
    reporter,
  }) {
    const rawValue = this.readEffectiveValue(worksheet, column, rowNumber, merges);
    const numericValue = toFiniteNumber(rawValue);

    if (numericValue !== null && numericValue > 0) {
      return numericValue;
    }

    const missing = isBlank(rawValue);
    reporter.record({
      type: missing ? "MISSING_INPUT_VALUE" : "INVALID_NUMERIC_INPUT",
      location: `${column}${rowNumber}`,
      field,
      message: missing
        ? `Input cell ${column}${rowNumber} is blank.`
        : `Input cell ${column}${rowNumber} must be a positive number, but contains "${displayText(rawValue)}".`,
      resolution: `The non-dropdown ${field} output cell will receive "ERROR". Freight Class is a dropdown and will be left blank.`,
    });

    return ERROR_VALUE;
  }

  valueOrError({ value, column, rowNumber, field, reporter }) {
    if (!isBlank(value)) return value;

    reporter.record({
      type: "MISSING_INPUT_VALUE",
      location: `${column}${rowNumber}`,
      field,
      message: `Input cell ${column}${rowNumber} is blank.`,
      resolution: `The non-dropdown ${field} output cell will receive "ERROR".`,
    });

    return ERROR_VALUE;
  }

  findSkidRangeForRow(rowIndex, merges) {
    const groupingColumnIndexes = this.config.skidGroupingColumns.map((column) =>
      XLSX.utils.decode_col(column),
    );

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

  readEffectiveValue(worksheet, column, rowNumber, merges) {
    const cell = this.readEffectiveCell(worksheet, column, rowNumber, merges);

    // .v preserves raw values such as Z. .w is only a formatted-text fallback.
    return cell?.v ?? cell?.w ?? "";
  }

  readEffectiveCell(worksheet, column, rowNumber, merges) {
    const address = `${column}${rowNumber}`;
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

    if (!merge) return directCell;

    return worksheet[XLSX.utils.encode_cell(merge.s)] ?? directCell;
  }
}

function displayText(value) {
  return String(value ?? "").trim();
}
