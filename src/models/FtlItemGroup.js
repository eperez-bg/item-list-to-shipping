/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FtlItemGroup.js                                                          ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ Holds all occurrences of one Item Code on one Truck # before calculations.║
 * ║ Grouping raw rows first prevents repeated items from producing multiple   ║
 * ║ output rows or losing quantity, pallet count, or combined weight.            ║
 * ║                                                                          ║
 * ║ FtlItemGroup — accumulates raw items and computes post-merge metrics.      ║
 * ║ addItem      — preserves source order for reference fields.                ║
 * ║ sumQuantity / sumWeight / calculateTotalValue — aggregate group metrics. ║
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
   * Appends source rows without calculating totals. Only after all duplicates
   * have joined do we compute combined quantity, direct weight sum, and value.
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
   * Parameters: none. Returns: sum of original G.W. (LBS) values or null.
   * FTL total weight is additive across source rows sharing Truck # and Item
   * Code; casepack quantity plays NO role in the calculation. Each item
   * contributes its original grossWeight (not LTL's allocatedWeight).
   * Missing/invalid weights fail the aggregate rather than becoming zero.
   */
  sumWeight() {
    const weights = this.items.map((item) => toFiniteNumber(item.grossWeight));
    if (weights.some((weight) => weight === null || weight <= 0)) return null;
    return weights.reduce((sum, weight) => sum + weight, 0);
  }

  /**
   * getCasepackPrice
   * Parameters: none. Returns: first item's nonnegative price or null.
   * Repeated items have one casepack price per combined output row, not a sum
   * of prices. The first source item's price is retained consistently; the
   * output factory warns if other matching rows have different prices.
   */
  getCasepackPrice() {
    const price = toFiniteNumber(this.representativeItem.casepackPrice);
    return price !== null && price >= 0 ? price : null;
  }

  /**
   * calculateTotalValue
   * Parameters: combinedQuantity (sum of target order casepack qty).
   * Returns: combined quantity × retained casepack price, or null if invalid.
   * The quantity is summed first but the price is never summed or multiplied
   * by the number of input rows.
   */
  calculateTotalValue(combinedQuantity) {
    const quantity = toFiniteNumber(combinedQuantity);
    const price = this.getCasepackPrice();
    if (quantity === null || quantity < 0 || price === null) return null;
    return quantity * price;
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
