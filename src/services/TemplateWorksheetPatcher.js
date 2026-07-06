/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: TemplateWorksheetPatcher.js                                          ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Safely writes output values into an in-memory copy of worksheet XML        ║
 * ║ while preserving styles and extending dropdown validation ranges.          ║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ Re-saving a workbook through a general library can lose template           ║
 * ║ features; patching the existing sheet XML preserves its styles and         ║
 * ║ validation metadata.                                                       ║
 * ║                                                                            ║
 * ║ MAIN CLASSES / FUNCTIONS                                                   ║
 * ║ • TemplateWorksheetPatcher — Re-saving a workbook through a general        ║
 * ║ library can lose template features; patching the existing sheet XML        ║
 * ║ preserves its styles and validation metadata.                              ║
 * ║ • createFilledWorkbookBlob — This is the final non-mutating write step     ║
 * ║ that turns logical output rows into a downloadable workbook copy.          ║
 * ║ • getStyleByColumn — Newly created cells should look like the template     ║
 * ║ instead of appearing unformatted.                                          ║
 * ║ • getOrCreateRow — Output can need more rows than the template already     ║
 * ║ contains.                                                                  ║
 * ║ • getOrCreateCell — Each mapped output field needs a cell even when the    ║
 * ║ template did not already contain that specific coordinate.                 ║
 * ║ • setCellValue — Old formulas/values must be cleared before writing a      ║
 * ║ new output value, and Excel needs different XML for numbers versus text.   ║
 * ║ • extendDataValidations — Generated rows must keep template dropdowns      ║
 * ║ even beyond the template's original preformatted range.                    ║
 * ║ • expandDimension — Excel worksheet bounds must include newly written      ║
 * ║ rows/columns so applications recognize them reliably.                      ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import * as XLSX from "xlsx";
import { extendSqrefThroughRow } from "../utils/a1";
import { createElementLike, elementChildren, firstDescendant, parseXml } from "../utils/xml";

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ CLASS: TemplateWorksheetPatcher                                            │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Template ZIP, worksheet path, layout, and completed output rows.
 *
 * Returns
 *   • A reusable safe worksheet patcher.
 *
 * Why / How
 *   Re-saving a workbook through a general library can lose template
 *   features; patching the existing sheet XML preserves its styles and
 *   validation metadata.
 *   It edits only the selected worksheet inside the template ZIP, reuses
 *   styles, writes values, extends validations, and returns a new Blob.
 */
export class TemplateWorksheetPatcher {
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: createFilledWorkbookBlob                                           │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • ZIP, XML path, TemplateLayout, output row models.
   *
   * Returns
   *   • Promise<Blob> for download.
   *
   * Why / How
   *   This is the final non-mutating write step that turns logical output rows
   *   into a downloadable workbook copy.
   *   It parses XML, clears/writes mapped cells, preserves styles, extends
   *   dropdown coverage, expands worksheet bounds, and regenerates the ZIP.
   */
  async createFilledWorkbookBlob({ zip, worksheetXmlPath, layout, outputRows }) {
    const worksheetFile = zip.file(worksheetXmlPath);
    if (!worksheetFile) {
      throw new Error(`Could not find ${worksheetXmlPath} inside the template workbook.`);
    }

    const worksheetXml = await worksheetFile.async("string");
    const xmlDocument = parseXml(worksheetXml, "output worksheet");
    const sheetData = firstDescendant(xmlDocument, "sheetData");

    if (!sheetData) {
      throw new Error("The template output worksheet has no sheetData element.");
    }

    const sourceStyles = this.getStyleByColumn(sheetData, layout);
    const lastOutputRow = layout.firstDataRow + outputRows.length - 1;
    const clearThroughRow = Math.max(layout.lastExistingDataRow, lastOutputRow);

    for (let rowNumber = layout.firstDataRow; rowNumber <= clearThroughRow; rowNumber += 1) {
      const templateRow = this.getOrCreateRow(xmlDocument, sheetData, rowNumber);
      const values = outputRows[rowNumber - layout.firstDataRow]?.toTemplateValues();

      Object.entries(layout.columnsByKey).forEach(([fieldKey, columnIndex]) => {
        const cellAddress = XLSX.utils.encode_cell({ c: columnIndex, r: rowNumber - 1 });
        const cell = this.getOrCreateCell(
          xmlDocument,
          templateRow,
          cellAddress,
          sourceStyles.get(columnIndex),
        );

        this.setCellValue(xmlDocument, cell, values?.[fieldKey] ?? "");
      });
    }

    this.extendDataValidations(xmlDocument, layout.firstDataRow, lastOutputRow);
    this.expandDimension(xmlDocument, lastOutputRow, layout.maxMappedColumnIndex);

    zip.file(worksheetXmlPath, new XMLSerializer().serializeToString(xmlDocument));

    return zip.generateAsync({
      type: "blob",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    });
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getStyleByColumn                                                   │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • sheetData XML node and layout.
   *
   * Returns
   *   • Map from column index to Excel style index.
   *
   * Why / How
   *   Newly created cells should look like the template instead of appearing
   *   unformatted.
   *   It finds the first styled existing data cell for each mapped column and
   *   caches its style index.
   */
  getStyleByColumn(sheetData, layout) {
    const styles = new Map();
    const rows = elementChildren(sheetData, "row");

    Object.values(layout.columnsByKey).forEach((columnIndex) => {
      if (styles.has(columnIndex)) return;

      for (const row of rows) {
        const rowNumber = Number(row.getAttribute("r"));
        if (rowNumber < layout.firstDataRow) continue;

        const matchingCell = elementChildren(row, "c").find((cell) => {
          const coordinate = XLSX.utils.decode_cell(cell.getAttribute("r"));
          return coordinate.c === columnIndex && cell.hasAttribute("s");
        });

        if (matchingCell) {
          styles.set(columnIndex, matchingCell.getAttribute("s"));
          break;
        }
      }
    });

    return styles;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getOrCreateRow                                                     │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • XML document, sheetData node, 1-based row number.
   *
   * Returns
   *   • Row XML element.
   *
   * Why / How
   *   Output can need more rows than the template already contains.
   *   It returns an existing row or inserts a new row in numerical order.
   */
  getOrCreateRow(xmlDocument, sheetData, rowNumber) {
    const existingRow = elementChildren(sheetData, "row").find(
      (row) => Number(row.getAttribute("r")) === rowNumber,
    );
    if (existingRow) return existingRow;

    const row = createElementLike(xmlDocument, sheetData, "row");
    row.setAttribute("r", String(rowNumber));

    const nextRow = elementChildren(sheetData, "row").find(
      (candidate) => Number(candidate.getAttribute("r")) > rowNumber,
    );

    sheetData.insertBefore(row, nextRow ?? null);
    return row;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getOrCreateCell                                                    │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • XML document, row node, A1 address, optional style index.
   *
   * Returns
   *   • Cell XML element.
   *
   * Why / How
   *   Each mapped output field needs a cell even when the template did not
   *   already contain that specific coordinate.
   *   It returns an existing cell or creates one in column order and applies
   *   the copied style.
   */
  getOrCreateCell(xmlDocument, row, cellAddress, styleIndex) {
    const existingCell = elementChildren(row, "c").find(
      (cell) => cell.getAttribute("r") === cellAddress,
    );
    if (existingCell) return existingCell;

    const cell = createElementLike(xmlDocument, row, "c");
    cell.setAttribute("r", cellAddress);
    if (styleIndex !== undefined) cell.setAttribute("s", styleIndex);

    const cellColumn = XLSX.utils.decode_cell(cellAddress).c;
    const nextCell = elementChildren(row, "c").find(
      (candidate) => XLSX.utils.decode_cell(candidate.getAttribute("r")).c > cellColumn,
    );

    row.insertBefore(cell, nextCell ?? null);
    return cell;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: setCellValue                                                       │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • XML document, cell element, output value.
   *
   * Returns
   *   • No returned value; mutates the cell XML.
   *
   * Why / How
   *   Old formulas/values must be cleared before writing a new output value,
   *   and Excel needs different XML for numbers versus text.
   *   It removes existing content, writes numeric values as v nodes, and
   *   writes strings as inlineStr while preserving edge whitespace.
   */
  setCellValue(xmlDocument, cell, value) {
    elementChildren(cell, "f").forEach((node) => node.remove());
    elementChildren(cell, "v").forEach((node) => node.remove());
    elementChildren(cell, "is").forEach((node) => node.remove());
    cell.removeAttribute("t");

    if (value === null || value === undefined || value === "") return;

    if (typeof value === "number" && Number.isFinite(value)) {
      const valueNode = createElementLike(xmlDocument, cell, "v");
      valueNode.textContent = String(value);
      cell.appendChild(valueNode);
      return;
    }

    const inlineString = createElementLike(xmlDocument, cell, "is");
    const text = createElementLike(xmlDocument, inlineString, "t");
    const stringValue = String(value);

    if (/^\s|\s$/.test(stringValue)) {
      text.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
    }

    text.textContent = stringValue;
    inlineString.appendChild(text);
    cell.setAttribute("t", "inlineStr");
    cell.appendChild(inlineString);
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: extendDataValidations                                              │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Worksheet XML, first data row, final output row.
   *
   * Returns
   *   • No returned value; mutates validation sqrefs.
   *
   * Why / How
   *   Generated rows must keep template dropdowns even beyond the template's
   *   original preformatted range.
   *   It expands each data-validation sqref that includes the first data row
   *   through the last output row.
   */
  extendDataValidations(xmlDocument, firstDataRow, lastOutputRow) {
    if (lastOutputRow < firstDataRow) return;

    Array.from(xmlDocument.getElementsByTagName("*") ?? [])
      .filter((node) => node.localName === "dataValidation")
      .forEach((validation) => {
        const sqref = validation.getAttribute("sqref");
        if (!sqref) return;

        validation.setAttribute(
          "sqref",
          extendSqrefThroughRow(sqref, firstDataRow, lastOutputRow),
        );
      });
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: expandDimension                                                    │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Worksheet XML, final row, last mapped column.
   *
   * Returns
   *   • No returned value; mutates the dimension range.
   *
   * Why / How
   *   Excel worksheet bounds must include newly written rows/columns so
   *   applications recognize them reliably.
   *   It decodes the current dimension, extends its end row/column, and writes
   *   it back.
   */
  expandDimension(xmlDocument, lastOutputRow, lastColumnIndex) {
    const dimension = firstDescendant(xmlDocument, "dimension");
    if (!dimension || lastOutputRow < 1) return;

    let range;
    try {
      range = XLSX.utils.decode_range(dimension.getAttribute("ref") ?? "A1");
    } catch {
      range = XLSX.utils.decode_range("A1");
    }

    range.e.r = Math.max(range.e.r, lastOutputRow - 1);
    range.e.c = Math.max(range.e.c, lastColumnIndex);
    dimension.setAttribute("ref", XLSX.utils.encode_range(range));
  }
}
