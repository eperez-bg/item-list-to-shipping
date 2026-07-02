import * as XLSX from "xlsx";
import { InputItem } from "../models/InputItem";
import { InputShipment } from "../models/InputShipment";
import { InputSkid } from "../models/InputSkid";
import { UserFacingError } from "../utils/errors";
import { ERROR_VALUE, ImportIssueReporter } from "../utils/importIssues";
import {
  getCellDisplayValue,
  isBlank,
  normalizeText,
  toFiniteNumber,
} from "../utils/text";

const ADDITIONAL_PALLET_ERROR_VALUE = "Error: Additional Pallet";
const ADDITIONAL_PALLET_CONTAINER_VALUE = "additional pallet";

export class InputWorkbookReader {
  constructor(inputConfig, shippingMode) {
    this.config = inputConfig;
    this.shippingMode = shippingMode;
    this.modeConfig = inputConfig.modes?.[shippingMode];

    if (!this.modeConfig) {
      throw new Error(
        `No input configuration exists for shipping mode "${shippingMode}".`,
      );
    }
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
    const weightGroups = new Map();
    let hasReachedData = false;

    for (
      let rowIndex = headerMap.headerRowIndex + 1;
      rowIndex <= range.e.r;
      rowIndex += 1
    ) {
      const rowNumber = rowIndex + 1;
      const container = this.readOptionalHeaderValue(
        worksheet,
        headerMap.columns.container,
        rowNumber,
        merges,
      );
      const isAdditionalPallet = this.isFtl() && this.isAdditionalPallet(container);

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

      // Blank rows and pre-data notes are ignored. A marked additional-pallet
      // row must stay in the output even though its regular item fields can be blank.
      if (isBlank(rawOldItemCode) && isBlank(rawQuantity) && !isAdditionalPallet) {
        continue;
      }
      if (!hasReachedData && isBlank(rawQuantity) && !isAdditionalPallet) {
        continue;
      }
      hasReachedData = true;

      const customerPo = this.readRequiredTextValue({
        worksheet,
        merges,
        columnIndex: headerMap.columns.customerPo,
        rowNumber,
        field: "Customer PO",
        reporter,
      });
      const oldItemCode = this.readRequiredTextValue({
        worksheet,
        merges,
        columnIndex: headerMap.columns.oldItemCode,
        rowNumber,
        field: "Old Item Code",
        reporter,
        isAdditionalPallet,
      });
      const quantity = this.readRequiredTextValue({
        worksheet,
        merges,
        columnIndex: headerMap.columns.quantity,
        rowNumber,
        field: "Quantity",
        reporter,
        isAdditionalPallet,
      });

      const targetStore = this.readEffectiveValue(
        worksheet,
        headerMap.columns.targetStore,
        rowNumber,
        merges,
      );
      const targetStoreSourceLocation = this.addressFor(
        headerMap.columns.targetStore,
        rowNumber,
      );

      if (isBlank(targetStore)) {
        reporter.record({
          type: "MISSING_INPUT_VALUE",
          location: targetStoreSourceLocation,
          field: "Target Store",
          message: `Input cell ${targetStoreSourceLocation} is blank.`,
          resolution:
            "Destination is a template dropdown, so it will be left blank. FTL pickup references will use ERROR for the missing Store #.",
        });
      }

      const pickupNumber = this.isLtl()
        ? this.readRequiredTextValue({
            worksheet,
            merges,
            columnIndex: headerMap.columns.pickupNumber,
            rowNumber,
            field: "Pickup Number",
            reporter,
          })
        : "";

      const truckNumber = this.isFtl()
        ? this.readRequiredTextValue({
            worksheet,
            merges,
            columnIndex: headerMap.columns.truckNumber,
            rowNumber,
            field: "Truck #",
            reporter,
          })
        : "";

      const pallet = this.isFtl()
        ? this.readPositiveNumberOrError({
            worksheet,
            merges,
            columnIndex: headerMap.columns.pallet,
            rowNumber,
            field: "Pallet",
            reporter,
          })
        : "";

      const chinaPo = this.isFtl()
        ? this.readRequiredTextValue({
            worksheet,
            merges,
            columnIndex: headerMap.columns.chinaPo,
            rowNumber,
            field: "China PO",
            reporter,
          })
        : "";

      const skidRange = this.findSkidRangeForRow(
        rowIndex,
        merges,
        headerMap.skidGroupingColumnIndexes,
      );
      const skidKey = `${skidRange.startRow}:${skidRange.endRow}`;
      const weightRange = this.findWeightRangeForRow(
        rowIndex,
        headerMap.columns.grossWeight,
        merges,
      );
      const weightKey = `${weightRange.startRow}:${weightRange.endRow}:${headerMap.columns.grossWeight}`;

      const item = new InputItem({
        sourceRow: rowNumber,
        customerPo,
        targetStore,
        targetStoreSourceLocation,
        oldItemCode,
        quantity,
        pickupNumber,
        truckNumber,
        pallet,
        chinaPo,
        container,
        isAdditionalPallet,
      });

      if (!groupedItems.has(skidKey)) {
        groupedItems.set(skidKey, { range: skidRange, items: [] });
      }
      groupedItems.get(skidKey).items.push(item);

      if (!weightGroups.has(weightKey)) {
        weightGroups.set(weightKey, { range: weightRange, items: [] });
      }
      weightGroups.get(weightKey).items.push(item);
    }

    if (groupedItems.size === 0) {
      this.throwCriticalInputError({
        type: "NO_INPUT_ITEMS_FOUND",
        location: `${sheetName}!${headerMap.headerRowNumber}:${headerMap.headerRowNumber}`,
        field: "Input item rows",
        message: "No item rows were found below the required input header row.",
        resolution:
          "Check that item rows contain values in both Old Item code and QTY, or are marked Additional Pallet in the container column.",
      });
    }

    this.allocateItemWeights({
      worksheet,
      merges,
      weightGroups,
      grossWeightColumnIndex: headerMap.columns.grossWeight,
      reporter,
    });

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

  isLtl() {
    return this.shippingMode === "LTL";
  }

  isFtl() {
    return this.shippingMode === "FTL";
  }

  isAdditionalPallet(containerValue) {
    return normalizeText(containerValue) === ADDITIONAL_PALLET_CONTAINER_VALUE;
  }

  findRequiredHeaderColumns(worksheet, range, sheetName) {
    const configuredHeaders = this.config.headers ?? {};
    const requiredFields = this.modeConfig.requiredFields ?? [];
    const optionalFields = this.modeConfig.optionalFields ?? [];
    const requiredHeaderTexts = requiredFields.map((field) => configuredHeaders[field]);

    if (requiredFields.length === 0 || requiredHeaderTexts.some(isBlank)) {
      this.throwCriticalInputError({
        type: "INPUT_HEADER_CONFIGURATION_ERROR",
        location: "Application configuration",
        field: "Input headers",
        message: `No complete required-header configuration exists for ${this.shippingMode}.`,
        resolution: "Add valid field names to input.modes in transformConfig.js.",
      });
    }

    const normalizedRequiredHeaders = requiredHeaderTexts.map(normalizeText);
    const duplicateConfiguredHeaders = normalizedRequiredHeaders.filter(
      (header, index) => normalizedRequiredHeaders.indexOf(header) !== index,
    );

    if (duplicateConfiguredHeaders.length > 0) {
      this.throwCriticalInputError({
        type: "DUPLICATE_INPUT_HEADER_CONFIGURATION",
        location: "Application configuration",
        field: "Input headers",
        message:
          "Each configured input header must be unique after case normalization. " +
          `Duplicate values: ${[...new Set(duplicateConfiguredHeaders)].join(", ")}.`,
        resolution: "Correct the headers object in transformConfig.js.",
      });
    }

    const trackedFields = [...new Set([...requiredFields, ...optionalFields])];
    const trackedHeaderKeys = new Set(
      trackedFields
        .map((field) => configuredHeaders[field])
        .filter((header) => !isBlank(header))
        .map(normalizeText),
    );
    const matchingColumnsByRow = new Map();

    for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
      const matchesForRow = new Map();

      for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
        const cell = worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })];
        const cellText = getCellDisplayValue(cell).trim();
        const headerKey = normalizeText(cellText);

        // Exact header matching is intentional, except capitalization does not matter.
        if (!trackedHeaderKeys.has(headerKey)) continue;
        if (!matchesForRow.has(headerKey)) matchesForRow.set(headerKey, []);
        matchesForRow.get(headerKey).push(columnIndex);
      }

      if (matchesForRow.size > 0) matchingColumnsByRow.set(rowIndex, matchesForRow);
    }

    const candidateHeaderRows = Array.from(matchingColumnsByRow.entries())
      .filter(([, matchesForRow]) =>
        requiredFields.every(
          (field) =>
            (matchesForRow.get(normalizeText(configuredHeaders[field])) ?? []).length === 1,
        ),
      )
      .map(([rowIndex, matchesForRow]) => ({ rowIndex, matchesForRow }));

    if (candidateHeaderRows.length === 0) {
      const foundHeaderLocations = this.describeFoundHeaderLocations(matchingColumnsByRow);

      this.throwCriticalInputError({
        type: "REQUIRED_INPUT_HEADERS_NOT_FOUND",
        location: sheetName,
        field: "Input headers",
        message:
          `Could not find one row containing every required case-insensitive exact ${this.shippingMode} input header. ` +
          `Required headers: ${requiredHeaderTexts.join(", ")}. ` +
          `Found matching header cells: ${foundHeaderLocations || "none"}.`,
        resolution:
          "Confirm the input workbook contains every required header with the same spelling and punctuation. Capitalization does not matter.",
      });
    }

    if (candidateHeaderRows.length > 1) {
      const rowNumbers = candidateHeaderRows.map(({ rowIndex }) => rowIndex + 1).join(", ");
      this.throwCriticalInputError({
        type: "AMBIGUOUS_INPUT_HEADER_ROW",
        location: sheetName,
        field: "Input headers",
        message: `More than one row contains all required input headers: rows ${rowNumbers}.`,
        resolution: "Keep only one complete input-header row in the uploaded worksheet.",
      });
    }

    const [{ rowIndex: headerRowIndex, matchesForRow }] = candidateHeaderRows;
    const columns = {};

    trackedFields.forEach((field) => {
      const configuredHeader = configuredHeaders[field];
      if (isBlank(configuredHeader)) return;

      const columnIndexes = matchesForRow.get(normalizeText(configuredHeader)) ?? [];

      if (columnIndexes.length === 1) {
        columns[field] = columnIndexes[0];
      }
    });

    const skidGroupingColumnIndexes = (this.modeConfig.skidGroupingFields ?? [])
      .map((field) => columns[field])
      .filter((columnIndex) => Number.isInteger(columnIndex));

    if (skidGroupingColumnIndexes.length === 0) {
      this.throwCriticalInputError({
        type: "SKID_GROUPING_CONFIGURATION_ERROR",
        location: "Application configuration",
        field: "Skid grouping headers",
        message: "None of the configured skidGroupingFields resolved to a required input header.",
        resolution: "Use field keys from input.headers, such as length, width, and height.",
      });
    }

    console.info("[Fuse Order Template Filler]", {
      type: "INPUT_HEADERS_RESOLVED",
      shippingMode: this.shippingMode,
      worksheet: sheetName,
      headerRow: headerRowIndex + 1,
      headers: Object.fromEntries(
        Object.entries(columns).map(([field, columnIndex]) => [
          field,
          { header: configuredHeaders[field], column: XLSX.utils.encode_col(columnIndex) },
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
          locations.push(`${headerText} at ${XLSX.utils.encode_col(columnIndex)}${rowIndex + 1}`);
        });
      });
    });
    return locations.join("; ");
  }

  allocateItemWeights({ worksheet, merges, weightGroups, grossWeightColumnIndex, reporter }) {
    weightGroups.forEach(({ range: weightRange, items }) => {
      const isAdditionalPalletGroup = items.length > 0 && items.every(
        (item) => item.isAdditionalPallet,
      );
      const totalWeight = this.readPositiveNumberOrError({
        worksheet,
        merges,
        columnIndex: grossWeightColumnIndex,
        rowNumber: weightRange.startRow,
        field: "Gross Weight",
        reporter,
        isAdditionalPallet: isAdditionalPalletGroup,
      });
      const grossWeightSourceLocation = this.addressFor(
        grossWeightColumnIndex,
        weightRange.startRow,
      );
      const allocatedWeight = isErrorValue(totalWeight)
        ? totalWeight
        : totalWeight / items.length;

      items.forEach((item) => {
        item.grossWeight = totalWeight;
        item.grossWeightSourceLocation = grossWeightSourceLocation;
        item.allocatedWeight = allocatedWeight;
      });

      if (weightRange.isMerged && items.length > 1) {
        console.info("[Fuse Order Template Filler]", {
          type: "MERGED_WEIGHT_SPLIT",
          sourceRows: `${weightRange.startRow}-${weightRange.endRow}`,
          totalWeight,
          itemCount: items.length,
          weightPerItem: allocatedWeight,
          itemSourceRows: items.map((item) => item.sourceRow),
        });
      }
    });
  }

  createSkid(worksheet, merges, skidRange, items, reporter, columns) {
    const sourceRow = skidRange.startRow;
    const isAdditionalPalletGroup = items.length > 0 && items.every(
      (item) => item.isAdditionalPallet,
    );
    const length = this.readPositiveNumberOrError({
      worksheet,
      merges,
      columnIndex: columns.length,
      rowNumber: sourceRow,
      field: "Length",
      reporter,
      isAdditionalPallet: isAdditionalPalletGroup,
    });
    const width = this.readPositiveNumberOrError({
      worksheet,
      merges,
      columnIndex: columns.width,
      rowNumber: sourceRow,
      field: "Width",
      reporter,
      isAdditionalPallet: isAdditionalPalletGroup,
    });
    const height = this.readPositiveNumberOrError({
      worksheet,
      merges,
      columnIndex: columns.height,
      rowNumber: sourceRow,
      field: "Height",
      reporter,
      isAdditionalPallet: isAdditionalPalletGroup,
    });

    items.forEach((item) => {
      item.assignDimensions({
        length,
        width,
        height,
        sourceStartRow: skidRange.startRow,
        sourceEndRow: skidRange.endRow,
      });
    });

    return new InputSkid({
      sourceStartRow: skidRange.startRow,
      sourceEndRow: skidRange.endRow,
      length,
      width,
      height,
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
    isAdditionalPallet = false,
  }) {
    const rawValue = this.readEffectiveValue(worksheet, columnIndex, rowNumber, merges);
    const numericValue = toFiniteNumber(rawValue);
    const location = this.addressFor(columnIndex, rowNumber);

    if (numericValue !== null && numericValue > 0) return numericValue;

    const missing = isBlank(rawValue);
    if (missing && isAdditionalPallet) {
      return this.recordAdditionalPalletValueMissing({ reporter, location, field });
    }

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

  readRequiredTextValue({
    worksheet,
    merges,
    columnIndex,
    rowNumber,
    field,
    reporter,
    isAdditionalPallet = false,
  }) {
    const value = this.readEffectiveValue(worksheet, columnIndex, rowNumber, merges);
    if (!isBlank(value)) return value;

    const location = this.addressFor(columnIndex, rowNumber);
    if (isAdditionalPallet) {
      return this.recordAdditionalPalletValueMissing({ reporter, location, field });
    }

    reporter.record({
      type: "MISSING_INPUT_VALUE",
      location,
      field,
      message: `Input cell ${location} is blank.`,
      resolution: `The non-dropdown ${field} output cell will receive "ERROR".`,
    });
    return ERROR_VALUE;
  }

  recordAdditionalPalletValueMissing({ reporter, location, field }) {
    reporter.record({
      type: "ADDITIONAL_PALLET_SOURCE_VALUE_MISSING",
      location,
      field,
      message:
        `Input cell ${location} is blank on a row whose container is marked "Additional Pallet".`,
      resolution:
        `The applicable non-dropdown output field will receive "${ADDITIONAL_PALLET_ERROR_VALUE}". Freight Class is a dropdown and will be left blank.`,
    });

    return ADDITIONAL_PALLET_ERROR_VALUE;
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
      .sort((left, right) => (right.e.r - right.s.r) - (left.e.r - left.s.r));

    const bestRange = matchingRanges[0];
    if (!bestRange) return { startRow: rowIndex + 1, endRow: rowIndex + 1 };

    return { startRow: bestRange.s.r + 1, endRow: bestRange.e.r + 1 };
  }

  findWeightRangeForRow(rowIndex, grossWeightColumnIndex, merges) {
    const matchingMerge = merges.find((merge) => {
      const vertical = merge.e.r > merge.s.r;
      const containsRow = merge.s.r <= rowIndex && rowIndex <= merge.e.r;
      const touchesGrossWeightColumn =
        merge.s.c <= grossWeightColumnIndex && grossWeightColumnIndex <= merge.e.c;
      return vertical && containsRow && touchesGrossWeightColumn;
    });

    if (!matchingMerge) {
      return { startRow: rowIndex + 1, endRow: rowIndex + 1, isMerged: false };
    }

    return {
      startRow: matchingMerge.s.r + 1,
      endRow: matchingMerge.e.r + 1,
      isMerged: true,
    };
  }

  readOptionalHeaderValue(worksheet, columnIndex, rowNumber, merges) {
    if (!Number.isInteger(columnIndex)) return "";
    return this.readEffectiveValue(worksheet, columnIndex, rowNumber, merges);
  }

  readEffectiveValue(worksheet, columnIndex, rowNumber, merges) {
    const cell = this.readEffectiveCell(worksheet, columnIndex, rowNumber, merges);
    return cell?.v ?? cell?.w ?? "";
  }

  readEffectiveCell(worksheet, columnIndex, rowNumber, merges) {
    const address = this.addressFor(columnIndex, rowNumber);
    const directCell = worksheet[address];
    if (!isBlank(getCellDisplayValue(directCell))) return directCell;

    const coordinate = XLSX.utils.decode_cell(address);
    const merge = merges.find(
      (candidate) =>
        candidate.s.c <= coordinate.c &&
        coordinate.c <= candidate.e.c &&
        candidate.s.r <= coordinate.r &&
        coordinate.r <= candidate.e.r,
    );

    return merge ? worksheet[XLSX.utils.encode_cell(merge.s)] ?? directCell : directCell;
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

function isErrorValue(value) {
  return value === ERROR_VALUE || value === ADDITIONAL_PALLET_ERROR_VALUE;
}
