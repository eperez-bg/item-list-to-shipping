import JSZip from "jszip";
import * as XLSX from "xlsx";
import { FreightClassCalculator } from "./FreightClassCalculator";
import { OutputRowFactory } from "./OutputRowFactory";
import { TemplateDropdownService } from "./TemplateDropdownService";
import { TemplateLayoutService } from "./TemplateLayoutService";
import { TemplateWorksheetPatcher } from "./TemplateWorksheetPatcher";
import { UserFacingError } from "../utils/errors";
import { ImportIssueReporter } from "../utils/importIssues";
import { elementChildren, parseXml } from "../utils/xml";

export class TemplateWorkbookService {
  constructor({ appConfig, templateFields }) {
    this.appConfig = appConfig;
    this.layoutService = new TemplateLayoutService(appConfig.template, templateFields);
    this.worksheetPatcher = new TemplateWorksheetPatcher();
  }

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
    const freightClassCalculator = new FreightClassCalculator(
      this.appConfig.freight,
      issueReporter,
    );
    const outputRowFactory = new OutputRowFactory({
      layout,
      dropdownService,
      freightClassCalculator,
      defaults: this.appConfig.defaults,
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

  async fetchTemplate() {
    let response;

    try {
      response = await fetch(this.appConfig.template.url, { cache: "no-store" });
    } catch {
      throw new UserFacingError(
        `Could not load the template at public${this.appConfig.template.url}.`,
      );
    }

    if (!response.ok) {
      throw new UserFacingError(
        `Could not load the template at public${this.appConfig.template.url}. Check that the filename and path are exact.`,
      );
    }

    return response.arrayBuffer();
  }

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

    const relationship = elementChildren(
      relationshipsXml.documentElement,
      "Relationship",
    ).find((node) => node.getAttribute("Id") === relationshipId);

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
