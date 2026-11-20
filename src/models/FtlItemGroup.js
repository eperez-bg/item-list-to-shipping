/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FtlItemGroup.js                                                          ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ Holds all occurrences of one Item Code on one Truck # before calculations.║
 * ║ Grouping raw rows first prevents repeated items from producing multiple   ║
 * ║ output rows or losing quantity, pallet count, or per-row weight.            ║
 * ║                                                                          ║
 * ║ FtlItemGroup — accumulates raw items and computes post-merge metrics.      ║
 * ║ addItem      — preserves source order for reference fields.                ║
 * ║ sumQuantity / sumWeight / largestDimension — aggregate original rows.      ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import { toFiniteNumber } from "../utils/text";

export class FtlItemGroup {
  /**
   * FtlItemGroup
   * Parameters: truck (Truckload), firstItem (InputItem).
   * Keeps the first matching source item as the representative. Its PO/IF and
   * catalog codes are used when the merged rows disagree on reference fields.
   */
  constructor({ truck, firstItem }) {
    this.truck = truck;
    this.items = [firstItem];
  }

  /**
   * addItem
   * Parameters: item (InputItem from the same truck and Item Code).
   * Appends source rows without performing calculations, so sums always use
   * each original quantity and G.W. rather than already-aggregated results.
   */
  addItem(item) {
    this.items.push(item);
  }

  /**
   * representativeItem
   * Parameters: none. Returns: first matching source item.
   * Reference values remain stable when duplicate rows differ in CustomerPO,
   * IF, or source formatting.
   */
  get representativeItem() {
    return this.items[0];
  }

  /**
   * sourceRows
   * Parameters: none. Returns: list of original Excel row numbers.
   * Provides traceable diagnostics when an aggregate is invalid.
   */
  get sourceRows() {
    return this.items.map((item) => item.sourceRow);
  }

  /**
   * isOrder
   * Parameters: none. Returns: boolean.
   * The merged item stays Order if it contains the very first raw item of
   * its truck, even if other Commodity items have been combined into it.
   */
  get isOrder() {
    return this.items.includes(this.truck.items[0]);
  }

  /**
   * palletCount
   * Parameters: none. Returns: number of original source rows.
   * One merged output item represents one pallet per raw item occurrence;
   * this intentionally ignores the old input 'Pallet' column.
   */
  get palletCount() {
    return this.items.length;
  }

  /**
   * sumQuantity
   * Parameters: none. Returns: sum of target-order casepack quantities or null.
   * Each source row contributes its full casepack quantity. Invalid values
   * fail the aggregate rather than being silently treated as zero.
   */
  sumQuantity() {
    const values = this.items.map((item) => toFiniteNumber(item.targetOrderCasepackQty));
    return values.some((value) => value === null || value < 0)
      ? null
      : values.reduce((sum, value) => sum + value, 0);
  }

  /**
   * sumWeight
   * Parameters: none. Returns: sum of (casepack qty × source G.W.) or null.
   * Multiplication happens for EVERY original row before summing. This avoids
   * the common error of multiplying merged quantity by only the first weight.
   */
  sumWeight() {
    let total = 0;
    for (const item of this.items) {
      const quantity = toFiniteNumber(item.targetOrderCasepackQty);
      const unitWeight = toFiniteNumber(item.grossWeight);
      if (quantity === null || quantity < 0 || unitWeight === null || unitWeight <= 0) {
        return null;
      }
      total += quantity * unitWeight;
    }
    return total;
  }

  /**
   * largestDimension
   * Parameters: field ('length', 'width', 'height'). Returns: maximum or null.
   * A combined item occupies the worst-case dimension from its source rows;
   * it is unsafe to copy just the first row's dimensions when they differ.
   */
  largestDimension(field) {
    const values = this.items.map((item) => toFiniteNumber(item[field]));
    return values.some((value) => value === null || value <= 0)
      ? null
      : Math.max(...values);
  }
}
