import * as XLSX from "xlsx";

export class TemplateLayout {
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

  addressFor(fieldKey, rowNumber) {
    const columnIndex = this.columnsByKey[fieldKey];

    if (columnIndex === undefined) {
      throw new Error(`The template does not contain a column for ${fieldKey}.`);
    }

    return XLSX.utils.encode_cell({ c: columnIndex, r: rowNumber - 1 });
  }

  get maxMappedColumnIndex() {
    return Math.max(...Object.values(this.columnsByKey));
  }
}
