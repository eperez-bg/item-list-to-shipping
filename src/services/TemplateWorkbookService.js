/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: TemplateWorkbookService.js                                           ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Orchestrates template loading, layout discovery, dropdown validation,      ║
 * ║ row creation, XML patching, and generation of a downloaded copy.           ║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ Template generation involves several independent concerns, so this         ║
 * ║ service coordinates them without placing workbook details in the UI.       ║
 * ║                                                                            ║
 * ║ MAIN CLASSES / FUNCTIONS                                                   ║
 * ║ • TemplateWorkbookService — Template generation involves several           ║
 * ║ independent concerns, so this service coordinates them without placing     ║
 * ║ workbook details in the UI.                                                ║
 * ║ • templateUrl — The UI needs the active template path for useful error     ║
 * ║ reporting without reaching into config internals.                          ║
 * ║ • outputFileName — The UI needs the active download name without knowing   ║
 * ║ the template config shape.                                                 ║
 * ║ • createFilledTemplate — A template must remain untouched while a filled   ║
 * ║ copy is generated in memory, including preserved dropdown metadata.        ║
 * ║ • createOutputRowFactory — Only row-building logic differs between LTL     ║
 * ║ and FTL, so mode selection should be isolated in one small factory         ║
 * ║ choice.                                                                    ║
 * ║ • fetchTemplate — The browser must load the template from the app's        ║
 * ║ public assets before it can create an isolated in-memory copy.             ║
 * ║ • findWorksheetXmlPath — SheetJS identifies a sheet by name while JSZip    ║
 * ║ patches a file path, so the relationship must be resolved correctly.       ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import JSZip from "jszip";
import * as XLSX from "xlsx";
import { FreightClassCalculator } from "./FreightClassCalculator";
import { FtlOutputRowFactory } from "./FtlOutputRowFactory";
import { OutputRowFactory } from "./OutputRowFactory";
import { TemplateDropdownService } from "./TemplateDropdownService";
import { TemplateLayoutService } from "./TemplateLayoutService";
import { TemplateWorksheetPatcher } from "./TemplateWorksheetPatcher";
import { UserFacingError } from "../utils/errors";
import { ImportIssueReporter } from "../utils/importIssues";
import { elementChildren, parseXml } from "../utils/xml";

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ CLASS: TemplateWorkbookService                                             │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Application config, active template config/fields, and shipping mode.
 *
 * Returns
 *   • A reusable generation service.
 *
 * Why / How
 *   Template generation involves several independent concerns, so this
 *   service coordinates them without placing workbook details in the UI.
 *   It loads the selected template, discovers layout, validates dropdowns,
 *   chooses the correct row factory, patches XML, and returns a Blob.
 */
export class TemplateWorkbookService {
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ CONSTRUCTOR: constructor                                                   │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Configs and active shipping mode.
   *
   * Returns
   *   • A configured TemplateWorkbookService.
   *
   * Why / How
   *   Each selected mode has a different template and field map, while shared
   *   services should be created once per active configuration.
   *   It stores configs and creates layout and worksheet patching
   *   collaborators.
   */
  constructor({ appConfig, templateConfig, templateFields, shippingMode }) {
    this.appConfig = appConfig;
    this.templateConfig = templateConfig;
    this.shippingMode = shippingMode;
    this.layoutService = new TemplateLayoutService(templateConfig, templateFields);
    this.worksheetPatcher = new TemplateWorksheetPatcher();
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ GETTER: templateUrl                                                        │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None.
   *
   * Returns
   *   • Active template URL.
   *
   * Why / How
   *   The UI needs the active template path for useful error reporting without
   *   reaching into config internals.
   *   It exposes the configured URL as a read-only getter.
   */
  get templateUrl() {
    return this.templateConfig.url;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ GETTER: outputFileName                                                     │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None.
   *
   * Returns
   *   • Active output file name.
   *
   * Why / How
   *   The UI needs the active download name without knowing the template
   *   config shape.
   *   It exposes the configured output filename.
   */
  get outputFileName() {
    return this.templateConfig.outputFileName;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: createFilledTemplate                                               │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • shipment: parsed InputShipment.
   *
   * Returns
   *   • Promise resolving to blob, rows, layout, and issues.
   *
   * Why / How
   *   A template must remain untouched while a filled copy is generated in
   *   memory, including preserved dropdown metadata.
   *   It loads template bytes, discovers layout/XML, builds valid rows,
   *   patches a ZIP copy, and returns the resulting Blob plus diagnostics.
   */
  async createFilledTemplate(shipment) {
    const issueReporter = new ImportIssueReporter(shipment.issues ?? []);
    const templateBuffer = await this.fetchTemplate();
    const workbook = XLSX.read(templateBuffer, { type: "array", cellDates: true });
    const layout = this.layoutService.findLayout(workbook);
    const zip = await JSZip.loadAsync(templateBuffer);
    const worksheetXmlPath = await this.findWorksheetXmlPath(zip, layout.sheetName);
    const worksheetXmlFile = zip.file(worksheetXmlPath);

    if (!worksheetXmlFile) {
      throw new UserFacingError(`Could not open output worksheet XML at ${worksheetXmlPath}.`);
    }

    const worksheetXml = await worksheetXmlFile.async("string");
    const dropdownService = new TemplateDropdownService({
      workbook,
      worksheetXml,
      destinationConfig: this.appConfig.destination,
      issueReporter,
    });
    const freightClassCalculator = new FreightClassCalculator(this.appConfig.freight, issueReporter);
    const outputRowFactory = this.createOutputRowFactory({
      layout,
      dropdownService,
      freightClassCalculator,
      issueReporter,
    });
    const outputRows = outputRowFactory.build(shipment);

    const blob = await this.worksheetPatcher.createFilledWorkbookBlob({
      zip,
      worksheetXmlPath,
      layout,
      outputRows,
    });

    return { blob, outputRows, layout, issues: issueReporter.issues };
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: createOutputRowFactory                                             │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Shared row-building services.
   *
   * Returns
   *   • Mode-appropriate output row factory.
   *
   * Why / How
   *   Only row-building logic differs between LTL and FTL, so mode selection
   *   should be isolated in one small factory choice.
   *   It builds shared dependencies and returns FtlOutputRowFactory for FTL or
   *   OutputRowFactory otherwise.
   */
  createOutputRowFactory({ layout, dropdownService, freightClassCalculator, issueReporter }) {
    const sharedOptions = {
      layout,
      dropdownService,
      freightClassCalculator,
      defaults: this.appConfig.defaults,
    };

    if (this.shippingMode === "FTL") {
      return new FtlOutputRowFactory({ ...sharedOptions, issueReporter });
    }

    return new OutputRowFactory(sharedOptions);
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: fetchTemplate                                                      │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None; uses configured URL.
   *
   * Returns
   *   • Promise<ArrayBuffer> or UserFacingError.
   *
   * Why / How
   *   The browser must load the template from the app's public assets before
   *   it can create an isolated in-memory copy.
   *   It fetches without cache, validates the HTTP result, and returns the raw
   *   ArrayBuffer.
   */
  async fetchTemplate() {
    let response;

    try {
      response = await fetch(this.templateConfig.url, { cache: "no-store" });
    } catch {
      throw new UserFacingError(`Could not load the template at ${this.templateConfig.url}.`);
    }

    if (!response.ok) {
      throw new UserFacingError(
        `Could not load the template at ${this.templateConfig.url}. Check that the filename and path are exact.`,
      );
    }

    return response.arrayBuffer();
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: findWorksheetXmlPath                                               │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Loaded JSZip workbook and selected sheet name.
   *
   * Returns
   *   • Promise resolving to worksheet XML path.
   *
   * Why / How
   *   SheetJS identifies a sheet by name while JSZip patches a file path, so
   *   the relationship must be resolved correctly.
   *   It reads workbook metadata and relationships XML, finds the named
   *   sheet's relationship target, and verifies the path exists in the ZIP.
   */
  async findWorksheetXmlPath(zip, sheetName) {
    const workbookXmlFile = zip.file("xl/workbook.xml");
    const relationshipsXmlFile = zip.file("xl/_rels/workbook.xml.rels");

    if (!workbookXmlFile || !relationshipsXmlFile) {
      throw new UserFacingError("The template does not look like a standard .xlsx workbook.");
    }

    const workbookXml = parseXml(await workbookXmlFile.async("string"), "workbook metadata");
    const relationshipsXml = parseXml(
      await relationshipsXmlFile.async("string"),
      "workbook relationships",
    );

    const sheet = Array.from(workbookXml.getElementsByTagName("*") ?? []).find(
      (node) => node.localName === "sheet" && node.getAttribute("name") === sheetName,
    );

    if (!sheet) {
      throw new UserFacingError(`Could not find worksheet "${sheetName}" in the template metadata.`);
    }

    const relationshipId =
      sheet.getAttribute("r:id") ??
      sheet.getAttributeNS(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "id",
      );

    const relationship = elementChildren(relationshipsXml.documentElement, "Relationship").find(
      (node) => node.getAttribute("Id") === relationshipId,
    );
    const target = relationship?.getAttribute("Target");

    if (!target) {
      throw new UserFacingError(`Could not find the file for worksheet "${sheetName}".`);
    }

    const path = new URL(target, "https://template.local/xl/workbook.xml")
      .pathname.replace(/^\//, "");

    if (!zip.file(path)) {
      throw new UserFacingError(`Could not find ${path} inside the template workbook.`);
    }

    return path;
  }
}
