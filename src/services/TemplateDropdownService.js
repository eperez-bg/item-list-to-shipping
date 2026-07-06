/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: TemplateDropdownService.js                                           ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Reads template data-validation dropdown lists and returns only values      ║
 * ║ that Excel will accept for dropdown-backed cells.                          ║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ Writing an arbitrary string into a dropdown-backed template cell can       ║
 * ║ make Excel reject or flag the workbook, so selection rules need one        ║
 * ║ central service.                                                           ║
 * ║                                                                            ║
 * ║ MAIN CLASSES / FUNCTIONS                                                   ║
 * ║ • TemplateDropdownService — Writing an arbitrary string into a             ║
 * ║ dropdown-backed template cell can make Excel reject or flag the            ║
 * ║ workbook, so selection rules need one central service.                     ║
 * ║ • resolveExactOptionOrBlank — Most dropdowns require an exact allowed      ║
 * ║ value, not a guess.                                                        ║
 * ║ • resolveOptionContainingOrBlank — Origin is configured as an address      ║
 * ║ fragment while the template dropdown may contain a longer full address.    ║
 * ║ • resolveDestinationOrBlank — Store values can have overlapping digits,    ║
 * ║ so destination matching must never use unsafe partial matches.             ║
 * ║ • getRawStoreNumber — Destination matching needs a single trustworthy      ║
 * ║ digit sequence while preserving meaningful leading zeroes.                 ║
 * ║ • pickDestinationMatch — Duplicate identical source-list entries should    ║
 * ║ not block an otherwise valid exact store selection.                        ║
 * ║ • getDestinationOptions — A newly generated output row may sit outside     ║
 * ║ the original validation sqref, so a destination fallback list is needed.   ║
 * ║ • getWorkbookDestinationOptions — Some template validation formulas        ║
 * ║ cannot be resolved, so workbook cells provide a last-resort exact store    ║
 * ║ list.                                                                      ║
 * ║ • findDestinationMatches — Exact digit comparison prevents 2077 from       ║
 * ║ accidentally matching 20770.                                               ║
 * ║ • uniqueOptions — Template lists can contain repeated labels, but          ║
 * ║ matching logic should treat duplicates as one choice.                      ║
 * ║ • extractStoreNumberFromOption — Destination resolution should compare     ║
 * ║ only the complete number following the configured Target Store prefix.     ║
 * ║ • getOptionsForCell — Other dropdown fields need the options attached to   ║
 * ║ their exact output cell.                                                   ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import * as XLSX from "xlsx";
import { cellIsInSqref } from "../utils/a1";
import { isBlank, normalizeText } from "../utils/text";
import { elementChildren, firstDescendant, parseXml } from "../utils/xml";

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ CLASS: TemplateDropdownService                                             │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Workbook, worksheet XML, destination configuration, and issue
 *   • reporter.
 *
 * Returns
 *   • A reusable template dropdown resolver.
 *
 * Why / How
 *   Writing an arbitrary string into a dropdown-backed template cell can
 *   make Excel reject or flag the workbook, so selection rules need one
 *   central service.
 *   It reads worksheet validation metadata, resolves list sources, matches
 *   requested values safely, and records problems as blanks when no valid
 *   option exists.
 */
export class TemplateDropdownService {
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ CONSTRUCTOR: constructor                                                   │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Workbook, output sheet XML, destination config, reporter.
   *
   * Returns
   *   • A configured TemplateDropdownService.
   *
   * Why / How
   *   Dropdown metadata should be parsed once rather than re-read for every
   *   output cell.
   *   It stores dependencies, parses validations, and initializes fallback
   *   caches.
   */
  constructor({ workbook, worksheetXml, destinationConfig, issueReporter }) {
    this.workbook = workbook;
    this.destinationConfig = destinationConfig;
    this.issueReporter = issueReporter;
    this.validations = this.readValidations(worksheetXml);

    this.destinationOptionsFallback = null;
    this.workbookDestinationOptions = null;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: resolveExactOptionOrBlank                                          │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Template cell address, requested value, optional context.
   *
   * Returns
   *   • Exact template option or empty string.
   *
   * Why / How
   *   Most dropdowns require an exact allowed value, not a guess.
   *   It rejects blanks/unreadable lists, normalizes candidate text, returns
   *   one exact option, and logs a safe blank on failure.
   */
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

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: resolveOptionContainingOrBlank                                     │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Cell address, search fragment, optional context.
   *
   * Returns
   *   • Full valid template option or empty string.
   *
   * Why / How
   *   Origin is configured as an address fragment while the template dropdown
   *   may contain a longer full address.
   *   It prefers an exact match, otherwise accepts one unambiguous containing
   *   option and logs blank for none/many.
   */
  resolveOptionContainingOrBlank(cellAddress, searchText, context = {}) {
    const field = context.field ?? "Dropdown";
    const sourceLocation = context.sourceLocation ?? "Application default";

    if (isBlank(searchText)) {
      this.issueReporter.record({
        type: "DROPDOWN_LEFT_BLANK",
        location: cellAddress,
        field,
        message: `${field} has no usable search value from ${sourceLocation}.`,
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
        message:
          `The template dropdown list could not be read while trying to select an option containing "${searchText}".`,
        resolution: "The template dropdown cell was left blank.",
      });

      return "";
    }

    const normalizedSearchText = normalizeText(searchText);

    // Prefer the exact dropdown option first.
    const exactMatch = options.find(
      (option) => normalizeText(option) === normalizedSearchText,
    );

    if (exactMatch !== undefined) {
      return exactMatch;
    }

    // For Origin, this selects the full valid dropdown label containing
    // "2354 Davis Ave". Example: "2354 Davis Ave, Hayward, CA 94545".
    const matches = this.uniqueOptions(
      options.filter((option) =>
        normalizeText(option).includes(normalizedSearchText),
      ),
    );

    if (matches.length === 1) {
      return matches[0];
    }

    this.issueReporter.record({
      type:
        matches.length === 0
          ? "DROPDOWN_OPTION_NOT_FOUND"
          : "AMBIGUOUS_DROPDOWN_OPTION",
      location: cellAddress,
      field,
      message:
        matches.length === 0
          ? `No ${field} dropdown option contains "${searchText}".`
          : `More than one ${field} dropdown option contains "${searchText}".`,
      resolution: "The template dropdown cell was left blank.",
    });

    return "";
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: resolveDestinationOrBlank                                          │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Destination cell address, raw store value, and source location.
   *
   * Returns
   *   • Exact Target Store dropdown option or empty string.
   *
   * Why / How
   *   Store values can have overlapping digits, so destination matching must
   *   never use unsafe partial matches.
   *   It tries exact digits first, then one leading-zero form, checking
   *   validation options and a workbook fallback.
   */
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
      Matching order:

      1. Try the input number exactly.
         Example: 2077 -> Target Store #2077

      2. Only if no exact result exists, try the same number with one leading zero.
         Example: 2077 -> Target Store #02077

      This never allows a partial-number match. 2077 will not match 20770.
    */
    const candidateStoreNumbers = [rawStoreNumber, `0${rawStoreNumber}`];

    const directOptions = this.getDestinationOptions(cellAddress);
    const workbookOptions = this.getWorkbookDestinationOptions();

    for (const candidateStoreNumber of candidateStoreNumbers) {
      const directMatches = this.findDestinationMatches(
        directOptions,
        candidateStoreNumber,
      );

      const directSelection = this.pickDestinationMatch(directMatches, {
        source: "data validation list",
        cellAddress,
        sourceLocation,
        rawStoreValue: storeValue,
        candidateStoreNumber,
      });

      if (directSelection) {
        return directSelection;
      }

      /*
        Fallback for templates where the destination validation formula cannot
        be fully resolved, or when the output row sits outside the original
        validation range.
      */
      const workbookMatches = this.findDestinationMatches(
        workbookOptions,
        candidateStoreNumber,
      );

      const workbookSelection = this.pickDestinationMatch(workbookMatches, {
        source: "template workbook search fallback",
        cellAddress,
        sourceLocation,
        rawStoreValue: storeValue,
        candidateStoreNumber,
      });

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

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getRawStoreNumber                                                  │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Raw Store # value.
   *
   * Returns
   *   • Store digits or empty string.
   *
   * Why / How
   *   Destination matching needs a single trustworthy digit sequence while
   *   preserving meaningful leading zeroes.
   *   It accepts numeric, direct numeric text, or one-number labels; multiple
   *   unrelated numbers are rejected.
   */
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

    // Accepts: 2077, 02077, and Excel-style values such as 2077.0.
    // Leading zeroes in the source are preserved.
    const directMatch = text.match(/^(\d+)(?:\.0+)?$/);

    if (directMatch) {
      return directMatch[1];
    }

    // Also accepts a label such as "Store 2077", but rejects values with
    // multiple unrelated number groups.
    const numberGroups = text.match(/\d+/g) ?? [];

    return numberGroups.length === 1 ? numberGroups[0] : "";
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: pickDestinationMatch                                               │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Exact matches and diagnostic context.
   *
   * Returns
   *   • Selected option or empty string.
   *
   * Why / How
   *   Duplicate identical source-list entries should not block an otherwise
   *   valid exact store selection.
   *   It returns the first exact match and emits a warning when multiple exact
   *   labels exist.
   */
  pickDestinationMatch(matches, context) {
    if (matches.length === 0) {
      return "";
    }

    // Duplicated labels in a source list should not make a valid exact store
    // selection fail. uniqueOptions removes identical labels first.
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

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getDestinationOptions                                              │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Destination cell address.
   *
   * Returns
   *   • Array of destination option strings.
   *
   * Why / How
   *   A newly generated output row may sit outside the original validation
   *   sqref, so a destination fallback list is needed.
   *   It first reads validations applying to the cell; if none, it caches the
   *   largest destination-like validation list in the template.
   */
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
      The sheet patcher can extend a validation down to new rows after the
      output is created. Here we still need its source list before that patch.
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
          "The output row was outside the original Destination validation range. The app used the template Destination list and will extend validation in the downloaded workbook.",
      });
    }

    return fallback;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getWorkbookDestinationOptions                                      │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None; reads this.workbook.
   *
   * Returns
   *   • Cached array of Target Store-like options.
   *
   * Why / How
   *   Some template validation formulas cannot be resolved, so workbook cells
   *   provide a last-resort exact store list.
   *   It scans workbook cell display values once and caches strings containing
   *   a valid store marker.
   */
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

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: findDestinationMatches                                             │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Option strings and candidate store number.
   *
   * Returns
   *   • Unique exact-matching option strings.
   *
   * Why / How
   *   Exact digit comparison prevents 2077 from accidentally matching 20770.
   *   It extracts each option's complete store number and keeps only exact
   *   equals matches.
   */
  findDestinationMatches(options, storeNumber) {
    return this.uniqueOptions(
      options.filter(
        (option) => this.extractStoreNumberFromOption(option) === storeNumber,
      ),
    );
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: uniqueOptions                                                      │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Array of option strings.
   *
   * Returns
   *   • Deduplicated option array.
   *
   * Why / How
   *   Template lists can contain repeated labels, but matching logic should
   *   treat duplicates as one choice.
   *   It normalizes each option and retains the first occurrence of each
   *   normalized form.
   */
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

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: extractStoreNumberFromOption                                       │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Option label text.
   *
   * Returns
   *   • Store number digits or empty string.
   *
   * Why / How
   *   Destination resolution should compare only the complete number following
   *   the configured Target Store prefix.
   *   It builds a safe regular expression from the prefix and extracts the
   *   following full digit sequence.
   */
  extractStoreNumberFromOption(option) {
    const prefix = escapeRegExp(
      normalizeText(this.destinationConfig.optionPrefix),
    ).replace(/\s+/g, "\\s*");

    /*
      Example with optionPrefix = "Target Store #":

      Target Store #2077  -> 2077
      Target Store #02077 -> 02077
      Target Store #20770 -> 20770
    */
    const expression = new RegExp(`${prefix}\\s*(\\d+)\\b`, "i");
    const match = normalizeText(option).match(expression);

    return match ? match[1] : "";
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getOptionsForCell                                                  │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Template cell address.
   *
   * Returns
   *   • Option strings or empty array.
   *
   * Why / How
   *   Other dropdown fields need the options attached to their exact output
   *   cell.
   *   It finds an applicable validation rule and resolves its formula,
   *   recording a readable issue if that fails.
   */
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

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: readValidations                                                    │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Output worksheet XML text.
   *
   * Returns
   *   • Array of validation metadata objects.
   *
   * Why / How
   *   The service needs a small normalized view of data-validation XML instead
   *   of repeatedly traversing raw DOM nodes.
   *   It parses XML and returns every validation's sqref and formula1 text.
   */
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

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: tryResolveFormulaToOptions                                         │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Validation formula text.
   *
   * Returns
   *   • Resolved options or empty array.
   *
   * Why / How
   *   Fallback scanning should not crash the workflow when one unrelated
   *   validation formula is unsupported.
   *   It calls the full resolver and converts failures to an empty option
   *   list.
   */
  tryResolveFormulaToOptions(formula) {
    try {
      return this.resolveFormulaToOptions(formula);
    } catch {
      return [];
    }
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: resolveFormulaToOptions                                            │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Raw validation formula.
   *
   * Returns
   *   • Array of allowed option strings or an error.
   *
   * Why / How
   *   Excel dropdowns can use inline strings, named ranges, or worksheet
   *   ranges, all of which need a uniform interface.
   *   It identifies formula type, resolves named ranges when needed, rejects
   *   unsafe dynamic formulas, and delegates range reading.
   */
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

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: readOptionsFromRange                                               │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Workbook range reference such as Sheet1!$A$1:$A$10.
   *
   * Returns
   *   • Array of option strings.
   *
   * Why / How
   *   A resolved dropdown range still needs to be converted into the actual
   *   visible option text.
   *   It parses sheet/range references, reads every nonblank cell in that
   *   range, and returns their display values.
   */
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

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: findLastUnquotedBang                                               │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Reference text.
   *
   * Returns
   *   • Index of the sheet/range separator or -1.
   *
   * Why / How
   *   Sheet names can be quoted and may contain punctuation, so splitting on
   *   the first ! is unsafe.
   *   It scans text while tracking quoted sheet names and returns the last
   *   unquoted exclamation index.
   */
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

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: escapeRegExp                                                     │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • value: arbitrary text.
 *
 * Returns
 *   • Regex-safe text.
 *
 * Why / How
 *   Configured text is inserted into a RegExp, so special characters must be
 *   escaped to avoid changing matching behavior.
 *   It prefixes regex metacharacters with backslashes.
 */
function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
