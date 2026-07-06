/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: TemplateLayout.js                                                    ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Stores the discovered output worksheet layout and converts logical field   ║
 * ║ names into exact Excel cell addresses.                                     ║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ Template columns may move, so services need an object that translates      ║
 * ║ stable logical field keys into discovered Excel positions.                 ║
 * ║                                                                            ║
 * ║ MAIN CLASSES / FUNCTIONS                                                   ║
 * ║ • TemplateLayout — Template columns may move, so services need an object   ║
 * ║ that translates stable logical field keys into discovered Excel            ║
 * ║ positions.                                                                 ║
 * ║ • addressFor — Dropdown lookup and XML writing both need the exact same    ║
 * ║ cell address for a logical field.                                          ║
 * ║ • maxMappedColumnIndex — The XML patcher needs the farthest mapped         ║
 * ║ column to expand the worksheet dimension correctly.                        ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import * as XLSX from "xlsx";

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ CLASS: TemplateLayout                                                      │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Layout data discovered from the workbook.
 *
 * Returns
 *   • A TemplateLayout used by output factories and XML patching.
 *
 * Why / How
 *   Template columns may move, so services need an object that translates
 *   stable logical field keys into discovered Excel positions.
 *   It stores the selected sheet, header/data rows, mapped field columns,
 *   and existing row boundary.
 */
export class TemplateLayout {
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ CONSTRUCTOR: constructor                                                   │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • sheetName, header/data row positions, columns, and existing row
   *   • boundary.
   *
   * Returns
   *   • A TemplateLayout instance.
   *
   * Why / How
   *   All template-writing services should use the same discovered layout
   *   rather than re-scan workbook headers.
   *   It stores the worksheet and column mapping metadata.
   */
  constructor({
    sheetName,
    headerRow,
    firstDataRow,
    columnsByKey,
    lastExistingDataRow,
  }) {
    this.sheetName = sheetName;
    this.headerRow = headerRow;
    this.firstDataRow = firstDataRow;
    this.columnsByKey = columnsByKey;
    this.lastExistingDataRow = lastExistingDataRow;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: addressFor                                                         │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • fieldKey and 1-based rowNumber.
   *
   * Returns
   *   • An Excel A1 address such as O12.
   *
   * Why / How
   *   Dropdown lookup and XML writing both need the exact same cell address
   *   for a logical field.
   *   It looks up the mapped column, fails loudly if missing, and encodes it
   *   as A1 notation.
   */
  addressFor(fieldKey, rowNumber) {
    const columnIndex = this.columnsByKey[fieldKey];

    if (columnIndex === undefined) {
      throw new Error(`The template does not contain a column for ${fieldKey}.`);
    }

    return XLSX.utils.encode_cell({ c: columnIndex, r: rowNumber - 1 });
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ GETTER: maxMappedColumnIndex                                               │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None.
   *
   * Returns
   *   • Largest mapped column index.
   *
   * Why / How
   *   The XML patcher needs the farthest mapped column to expand the worksheet
   *   dimension correctly.
   *   It finds the maximum mapped numeric column index.
   */
  get maxMappedColumnIndex() {
    return Math.max(...Object.values(this.columnsByKey));
  }
}
