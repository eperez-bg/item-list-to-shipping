/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FtlCatalogService.js                                                     ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ Loads the static FTL catalog (BEELINEGROUPFX_L3.xlsx), discovers its      ║
 * ║ columns by header name, and indexes items for quick case-insensitive      ║
 * ║ input Item Code lookups using the catalog Manufacturer Part ID.           ║
 * ║ Header names are read from Excel row 11. FTL only.                       ║
 * ║                                                                          ║
 * ║ load — fetch and parse the public catalog; fromArrayBuffer — index it;    ║
 * ║ lookup — find matching catalog details by the input Item Code.            ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import * as XLSX from "xlsx";
import { UserFacingError } from "../utils/errors";
import { getCellDisplayValue, isBlank, normalizeText } from "../utils/text";

export class FtlCatalogService {
  /**
   * FtlCatalogService
   * Parameters: config with url, headerRowNumber (Excel row 11), and headers.
   * Separates catalog loading and indexing from both template patching and
   * FTL row generation so each has a single clear responsibility.
   */
  constructor(config) {
    this.config = config;
    this.recordsByManufacturerPartId = new Map();
  }

  /**
   * load
   * Parameters: none. Returns: Promise<FtlCatalogService>.
   * Reads the catalog from the app's public assets. The user must supply the
   * original .xlsx at the configured URL; no catalog is silently invented.
   */
  async load() {
    let response;
    try {
      response = await fetch(this.config.url);
    } catch (error) {
      throw new UserFacingError(
        `Could not fetch the FTL catalog at ${this.config.url}: ${error.message}`,
      );
    }
    if (!response.ok) {
      throw new UserFacingError(
        `FTL catalog was not found (${response.status}). Add BEELINEGROUPFX_L3.xlsx ` +
          `under public/catalogs/ and redeploy; expected URL: ${this.config.url}`,
      );
    }

    try {
      return this.fromArrayBuffer(await response.arrayBuffer());
    } catch (error) {
      if (error instanceof UserFacingError) throw error;
      throw new UserFacingError(`Could not parse the FTL catalog: ${error.message}`);
    }
  }

  /**
   * fromArrayBuffer
   * Parameters: catalog workbook bytes. Returns: this (with lookup index).
   * Checks row 11 of each worksheet for all three required catalog headers,
   * then indexes Manufacturer Part ID so source Item Code lookups are constant-time.
   */
  fromArrayBuffer(arrayBuffer) {
    const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: false });
    const expected = this.config.headers;
    const keys = ["manufacturerPartId", "description", "supplierPartId"];
    const columnMap = keys.map((key) => normalizeText(expected[key]));
    this.recordsByManufacturerPartId.clear();
    const headerRowNumber = this.config.headerRowNumber ?? 11;
    if (!Number.isSafeInteger(headerRowNumber) || headerRowNumber < 1) {
      throw new UserFacingError(
        `Invalid FTL catalog header row: ${headerRowNumber}. Set headerRowNumber to a positive Excel row number.`,
      );
    }

    let foundAnyHeader = false;
    const inspectedSheetNames = [];

    for (const sheetName of workbook.SheetNames) {
      const worksheet = workbook.Sheets[sheetName];
      if (!worksheet?.["!ref"]) continue;
      const range = XLSX.utils.decode_range(worksheet["!ref"]);
      inspectedSheetNames.push(sheetName);
      const headerRow = this.findHeaderRow(
        worksheet,
        range,
        columnMap,
        headerRowNumber,
      );
      if (!headerRow) continue;
      foundAnyHeader = true;
      console.info("[Fuse Order Template Filler]", {
        type: "FTL_CATALOG_HEADERS_FOUND",
        sheetName,
        headerRow: headerRowNumber,
        columns: Object.fromEntries(headerRow.columns),
      });

      for (let row = headerRow.row + 1; row <= range.e.r; row += 1) {
        const values = keys.map((key) => {
          const column = headerRow.columns.get(normalizeText(expected[key]));
          return getCellDisplayValue(worksheet[XLSX.utils.encode_cell({ r: row, c: column })]).trim();
        });
        const [manufacturerPartId, description, supplierPartId] = values;
        if (isBlank(manufacturerPartId)) continue;
        const lookupKey = normalizeText(manufacturerPartId);
        const record = {
          manufacturerPartId,
          description,
          supplierPartId,
          catalogLocation: `${sheetName}!${row + 1}`,
        };
        if (this.recordsByManufacturerPartId.has(lookupKey)) {
          const prior = this.recordsByManufacturerPartId.get(lookupKey);
          console.warn("[Fuse Order Template Filler]", {
            type: "DUPLICATE_FTL_CATALOG_MANUFACTURER_PART_ID",
            manufacturerPartId,
            location: record.catalogLocation,
            keptLocation: prior.catalogLocation,
            message: "Catalog has more than one row for this Manufacturer Part ID; keeping the first.",
          });
          continue;
        }
        this.recordsByManufacturerPartId.set(lookupKey, record);
      }
    }

    if (!foundAnyHeader) {
      throw new UserFacingError(
        `The FTL catalog does not contain all required headers on Excel row ${headerRowNumber}: ` +
        `${keys.map((key) => expected[key]).join(", ")}. ` +
        `Worksheets checked: ${inspectedSheetNames.join(", ") || "none"}.`,
      );
    }
    if (this.recordsByManufacturerPartId.size === 0) {
      throw new UserFacingError("The FTL catalog headers were found, but no Manufacturer Part ID records were present.");
    }

    console.info("[Fuse Order Template Filler]", {
      type: "FTL_CATALOG_LOADED",
      url: this.config.url,
      itemCount: this.recordsByManufacturerPartId.size,
    });
    return this;
  }

  /**
   * findHeaderRow
   * Parameters: worksheet, used range, normalized header names.
   * Returns: {row, columns} or null.
   * Uses the configured Excel header row (11) and finds columns by exact,
   * case-insensitive labels. Earlier title/notes rows cannot be mistaken for
   * column headers, and column order may still change safely.
   */
  findHeaderRow(worksheet, range, normalizedNames, headerRowNumber) {
    // Excel row numbers begin at 1; SheetJS row indices begin at 0.
    const row = headerRowNumber - 1;
    if (row < range.s.r || row > range.e.r) return null;

    const required = new Set(normalizedNames);
    const matches = new Map();
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const raw = getCellDisplayValue(
        worksheet[XLSX.utils.encode_cell({ r: row, c: col })],
      );
      const normalized = normalizeText(raw);
      if (required.has(normalized)) matches.set(normalized, col);
    }

    return normalizedNames.every((name) => matches.has(name))
      ? { row, columns: matches }
      : null;
  }

  /**
   * lookup
   * Parameters: input Item Code. Returns: matched catalog record or null.
   * The input Item Code matches catalog Manufacturer Part ID using
   * case-insensitive, whitespace-normalized comparison. The caller logs
   * an error when no matching catalog record exists.
   */
  lookup(itemCode) {
    return this.recordsByManufacturerPartId.get(normalizeText(itemCode)) ?? null;
  }
}
