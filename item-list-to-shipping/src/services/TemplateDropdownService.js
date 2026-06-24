import * as XLSX from "xlsx";
import { cellIsInSqref } from "../utils/a1";
import { isBlank, normalizeText } from "../utils/text";
import { elementChildren, firstDescendant, parseXml } from "../utils/xml";

export class TemplateDropdownService {
  constructor({ workbook, worksheetXml, destinationConfig, issueReporter }) {
    this.workbook = workbook;
    this.destinationConfig = destinationConfig;
    this.issueReporter = issueReporter;
    this.validations = this.readValidations(worksheetXml);

    this.destinationOptionsFallback = null;
    this.workbookDestinationOptions = null;
  }

  resolveExactOptionOrBlank(cellAddress, desiredValue, context = {}) {
    const field = context.field ?? "Dropdown";
    const sourceLocation = context.sourceLocation ?? "Application default";

    if (isBlank(desiredValue)) {
      this.issueReporter.record({
        type: "DROPDOWN_LEFT_BLANK",
        location: cellAddress,
        field,
        message: `${field} has no usable source value from ${sourceLocation}.`,
        resolution: "The template dropdown cell was left blank.",
      });

      return "";
    }

    const options = this.getOptionsForCell(cellAddress);

    if (options.length === 0) {
      this.issueReporter.record({
        type: "TEMPLATE_DROPDOWN_UNREADABLE",
        location: cellAddress,
        field,
        message: `The template dropdown list could not be read while trying to select "${desiredValue}".`,
        resolution: "The template dropdown cell was left blank.",
      });

      return "";
    }

    const normalizedDesired = normalizeText(desiredValue);

    const match = options.find(
      (option) => normalizeText(option) === normalizedDesired,
    );

    if (match === undefined) {
      this.issueReporter.record({
        type: "DROPDOWN_OPTION_NOT_FOUND",
        location: cellAddress,
        field,
        message: `"${desiredValue}" is not an allowed dropdown value.`,
        resolution: "The template dropdown cell was left blank.",
      });

      return "";
    }

    return match;
  }

  resolveDestinationOrBlank(cellAddress, storeValue, sourceLocation) {
    if (isBlank(storeValue)) {
      this.issueReporter.record({
        type: "DROPDOWN_LEFT_BLANK",
        location: cellAddress,
        field: "Destination",
        message: `There is no target-store value at ${sourceLocation}.`,
        resolution: "Destination is a template dropdown, so it was left blank.",
      });

      return "";
    }

    const rawStoreNumber = this.getRawStoreNumber(storeValue);

    if (isBlank(rawStoreNumber)) {
      this.issueReporter.record({
        type: "INVALID_DESTINATION_INPUT",
        location: sourceLocation,
        field: "Target Store",
        message: `"${storeValue}" cannot be converted to one target-store number.`,
        resolution: "Destination is a template dropdown, so it was left blank.",
      });

      return "";
    }

    /*
      Required matching order:

      1. Try the exact store number from input.
         Example: 2077 -> Target Store #2077

      2. Only if no exact result exists, prepend one zero.
         Example: 2077 -> Target Store #02077

      Similar values such as 20770, 12077, or 207700 will not match because
      the extracted store number must equal the candidate exactly.
    */
    const candidateStoreNumbers = [
      rawStoreNumber,
      `0${rawStoreNumber}`,
    ];

    const directOptions = this.getDestinationOptions(cellAddress);
    const workbookOptions = this.getWorkbookDestinationOptions();

    for (const candidateStoreNumber of candidateStoreNumbers) {
      const directMatches = this.findDestinationMatches(
        directOptions,
        candidateStoreNumber,
      );

      const directSelection = this.pickDestinationMatch(
        directMatches,
        {
          source: "data validation list",
          cellAddress,
          sourceLocation,
          rawStoreValue: storeValue,
          candidateStoreNumber,
        },
      );

      if (directSelection) {
        return directSelection;
      }

      /*
        Fallback for cases where the template's validation formula cannot be
        fully read, or the output row is outside the template's original
        validation range. This searches workbook cells containing labels like:
        Target Store #2077
      */
      const workbookMatches = this.findDestinationMatches(
        workbookOptions,
        candidateStoreNumber,
      );

      const workbookSelection = this.pickDestinationMatch(
        workbookMatches,
        {
          source: "template workbook search fallback",
          cellAddress,
          sourceLocation,
          rawStoreValue: storeValue,
          candidateStoreNumber,
        },
      );

      if (workbookSelection) {
        console.warn("[Fuse Order Template Filler]", {
          type: "DESTINATION_WORKBOOK_SEARCH_FALLBACK",
          location: cellAddress,
          sourceLocation,
          rawStoreValue: storeValue,
          candidateStoreNumber,
          selectedOption: workbookSelection,
          message:
            "The exact destination option was found by searching the template workbook.",
        });

        return workbookSelection;
      }
    }

    this.issueReporter.record({
      type: "DESTINATION_OPTION_NOT_FOUND",
      location: cellAddress,
      field: "Destination",
      message:
        `No exact Destination option was found for store "${rawStoreNumber}" ` +
        `or zero-prefixed store "0${rawStoreNumber}" from ${sourceLocation}.`,
      resolution: "Destination is a template dropdown, so it was left blank.",
    });

    console.error("[Fuse Order Template Filler] Destination not found", {
      type: "DESTINATION_OPTION_NOT_FOUND",
      location: cellAddress,
      sourceLocation,
      rawStoreValue: storeValue,
      exactStoreNumberTried: rawStoreNumber,
      zeroPrefixedStoreNumberTried: `0${rawStoreNumber}`,
      directDestinationOptions: directOptions,
      workbookDestinationOptions: workbookOptions,
    });

    return "";
  }

  getRawStoreNumber(value) {
    if (isBlank(value)) {
      return "";
    }

    if (typeof value === "number") {
      if (!Number.isSafeInteger(value) || value < 0) {
        return "";
      }

      return String(value);
    }

    const text = String(value)
      .normalize("NFKC")
      .replace(/\u00a0/g, " ")
      .replace(/,/g, "")
      .trim();

    /*
      Accepts:
      2077
      02077
      2077.0

      Keeps leading zeroes when they are present in the source file.
    */
    const directMatch = text.match(/^(\d+)(?:\.0+)?$/);

    if (directMatch) {
      return directMatch[1];
    }

    /*
      Also accepts a source value such as "Store 2077", but rejects source
      values that contain multiple unrelated numbers.
    */
    const numberGroups = text.match(/\d+/g) ?? [];

    return numberGroups.length === 1 ? numberGroups[0] : "";
  }

  pickDestinationMatch(matches, context) {
    if (matches.length === 0) {
      return "";
    }

    /*
      Do not leave the cell blank merely because the same store number appears
      more than once in the source list. Choose the first valid exact option.
    */
    if (matches.length > 1) {
      console.warn("[Fuse Order Template Filler]", {
        type: "MULTIPLE_EXACT_DESTINATION_OPTIONS",
        message:
          "More than one exact Destination option was found. The first valid option was selected.",
        ...context,
        matches,
        selectedOption: matches[0],
      });
    }

    return matches[0];
  }

  getDestinationOptions(cellAddress) {
    const directOptions = this.uniqueOptions(
      this.validations
        .filter((validation) => cellIsInSqref(cellAddress, validation.sqref))
        .flatMap((validation) =>
          this.tryResolveFormulaToOptions(validation.formula),
        ),
    );

    const directDestinationOptions = directOptions.filter(
      (option) => this.extractStoreNumberFromOption(option) !== "",
    );

    if (directDestinationOptions.length > 0) {
      return directDestinationOptions;
    }

    /*
      If the current output row is beyond the original validation range, use
      the biggest validation list in the template that contains Target Store
      values. The worksheet patcher later expands the validation range.
    */
    if (this.destinationOptionsFallback) {
      return this.destinationOptionsFallback;
    }

    const candidateLists = this.validations
      .map((validation) =>
        this.uniqueOptions(
          this.tryResolveFormulaToOptions(validation.formula).filter(
            (option) => this.extractStoreNumberFromOption(option) !== "",
          ),
        ),
      )
      .filter((options) => options.length > 0);

    const fallback =
      candidateLists.sort((left, right) => right.length - left.length)[0] ??
      [];

    this.destinationOptionsFallback = fallback;

    if (fallback.length > 0) {
      console.warn("[Fuse Order Template Filler]", {
        type: "DESTINATION_VALIDATION_RANGE_FALLBACK",
        location: cellAddress,
        message:
          "The output row was outside the original Destination validation range. The app used the template's Destination list and will extend validation in the downloaded workbook.",
      });
    }

    return fallback;
  }

  getWorkbookDestinationOptions() {
    if (this.workbookDestinationOptions) {
      return this.workbookDestinationOptions;
    }

    const options = [];

    Object.values(this.workbook.Sheets).forEach((worksheet) => {
      Object.keys(worksheet).forEach((address) => {
        if (address.startsWith("!")) {
          return;
        }

        const cell = worksheet[address];
        const value = cell?.w ?? cell?.v;

        if (isBlank(value)) {
          return;
        }

        const option = String(value);

        if (this.extractStoreNumberFromOption(option) !== "") {
          options.push(option);
        }
      });
    });

    this.workbookDestinationOptions = this.uniqueOptions(options);

    return this.workbookDestinationOptions;
  }

  findDestinationMatches(options, storeNumber) {
    return this.uniqueOptions(
      options.filter(
        (option) => this.extractStoreNumberFromOption(option) === storeNumber,
      ),
    );
  }

  uniqueOptions(options) {
    const seen = new Set();

    return options.filter((option) => {
      const key = normalizeText(option);

      if (!key || seen.has(key)) {
        return false;
      }

      seen.add(key);

      return true;
    });
  }

  extractStoreNumberFromOption(option) {
    const prefix = escapeRegExp(
      normalizeText(this.destinationConfig.optionPrefix),
    ).replace(/\s+/g, "\\s*");

    /*
      Example with optionPrefix = "Target Store #":

      Target Store #2077  -> 2077
      Target Store #02077 -> 02077
      Target Store #20770 -> 20770

      This does not transform, pad, trim, or partially match the number.
    */
    const expression = new RegExp(`${prefix}\\s*(\\d+)\\b`, "i");
    const match = normalizeText(option).match(expression);

    return match ? match[1] : "";
  }

  getOptionsForCell(cellAddress) {
    const validation = this.validations.find((candidate) =>
      cellIsInSqref(cellAddress, candidate.sqref),
    );

    if (!validation) {
      return [];
    }

    try {
      return this.resolveFormulaToOptions(validation.formula);
    } catch (error) {
      this.issueReporter.record({
        type: "TEMPLATE_DROPDOWN_UNREADABLE",
        location: cellAddress,
        field: "Template dropdown",
        message: error.message || "The dropdown source could not be resolved.",
        resolution: "The template dropdown cell was left blank.",
      });

      return [];
    }
  }

  readValidations(worksheetXml) {
    const xmlDocument = parseXml(worksheetXml, "worksheet validation rules");

    return Array.from(xmlDocument.getElementsByTagName("*") ?? [])
      .filter((node) => node.localName === "dataValidation")
      .map((node) => {
        const formulaNode =
          elementChildren(node, "formula1")[0] ??
          firstDescendant(node, "formula1");

        return {
          sqref: node.getAttribute("sqref") ?? "",
          formula: formulaNode?.textContent?.trim() ?? "",
        };
      })
      .filter((validation) => validation.sqref && validation.formula);
  }

  tryResolveFormulaToOptions(formula) {
    try {
      return this.resolveFormulaToOptions(formula);
    } catch {
      return [];
    }
  }

  resolveFormulaToOptions(formula) {
    const cleanedFormula = String(formula).trim();

    if (/^".*"$/.test(cleanedFormula)) {
      return cleanedFormula
        .slice(1, -1)
        .split(",")
        .map((value) => value.trim())
        .filter(Boolean);
    }

    const formulaReference = cleanedFormula.replace(/^=/, "").trim();

    const namedRange = this.workbook.Workbook?.Names?.find(
      (name) =>
        normalizeText(name.Name) === normalizeText(formulaReference),
    );

    const rangeReference = namedRange?.Ref ?? formulaReference;

    if (/\b(indirect|offset)\s*\(/i.test(rangeReference)) {
      throw new Error(
        `The template uses an unsupported INDIRECT or OFFSET dropdown formula (${formula}).`,
      );
    }

    return this.readOptionsFromRange(rangeReference);
  }

  readOptionsFromRange(rangeReference) {
    const withoutEquals = String(rangeReference)
      .replace(/^=/, "")
      .trim();

    const bangIndex = this.findLastUnquotedBang(withoutEquals);

    if (bangIndex < 0) {
      throw new Error(`Could not resolve dropdown source "${rangeReference}".`);
    }

    const rawSheetName = withoutEquals.slice(0, bangIndex).trim();

    const sheetName = rawSheetName
      .replace(/^'/, "")
      .replace(/'$/, "")
      .replace(/''/g, "'");

    const rawRange = withoutEquals
      .slice(bangIndex + 1)
      .replace(/\$/g, "")
      .trim();

    const worksheet = this.workbook.Sheets[sheetName];

    if (!worksheet) {
      throw new Error(
        `The dropdown source worksheet "${sheetName}" does not exist.`,
      );
    }

    let range;

    try {
      range = XLSX.utils.decode_range(rawRange);
    } catch {
      throw new Error(
        `The dropdown source range "${rangeReference}" is invalid.`,
      );
    }

    const values = [];

    for (let row = range.s.r; row <= range.e.r; row += 1) {
      for (let column = range.s.c; column <= range.e.c; column += 1) {
        const cell = worksheet[
          XLSX.utils.encode_cell({ r: row, c: column })
        ];

        const value = cell?.w ?? cell?.v;

        if (!isBlank(value)) {
          values.push(String(value));
        }
      }
    }

    return values;
  }

  findLastUnquotedBang(value) {
    let insideQuote = false;
    let lastBangIndex = -1;

    for (let index = 0; index < value.length; index += 1) {
      if (value[index] === "'") {
        if (insideQuote && value[index + 1] === "'") {
          index += 1;
          continue;
        }

        insideQuote = !insideQuote;
      }

      if (value[index] === "!" && !insideQuote) {
        lastBangIndex = index;
      }
    }

    return lastBangIndex;
  }
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}