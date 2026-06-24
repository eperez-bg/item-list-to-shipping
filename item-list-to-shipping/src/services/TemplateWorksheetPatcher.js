import * as XLSX from "xlsx";
import { extendSqrefThroughRow } from "../utils/a1";
import { createElementLike, elementChildren, firstDescendant, parseXml } from "../utils/xml";

export class TemplateWorksheetPatcher {
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
