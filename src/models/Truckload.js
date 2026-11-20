/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: Truckload.js                                                         ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Represents all items sharing one Truck # so FTL can identify its Order     ║
 * ║ row and find its pallet total.                                             ║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ FTL output groups by Truck # rather than PO, so this model centralizes     ║
 * ║ the Truck # Order/Commodity and pallet-count rules.                        ║
 * ║                                                                            ║
 * ║ MAIN CLASSES / FUNCTIONS                                                   ║
 * ║ • Truckload — FTL output groups by Truck # rather than PO, so this model   ║
 * ║ centralizes the Truck # Order/Commodity and pallet-count rules.            ║
 * ║ • addItem — Truck groups are built while the reader walks source rows,     ║
 * ║ so members must be appended consistently.                                  ║
 * ║ • itemCount — Callers need a readable truck member count.                  ║
 * ║ • isFirstItem — The first source item for a Truck # is the FTL Order       ║
 * ║ row.                                                                       ║
 * ║ • getGreatestPalletCount — The FTL Order row must show the truck's         ║
 * ║ maximum Pallet number, even when that number appears later in the input.   ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import { ERROR_VALUE } from "../utils/importIssues";
import { normalizeText, toFiniteNumber } from "../utils/text";
import { FtlItemGroup } from "./FtlItemGroup";

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ CLASS: Truckload                                                           │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Constructor options for a truck group.
 *
 * Returns
 *   • Instances used by FTL output and preview generation.
 *
 * Why / How
 *   FTL output groups by Truck # rather than PO, so this model centralizes
 *   the Truck # Order/Commodity and pallet-count rules.
 *   It stores truck members in source order and calculates the greatest
 *   positive Pallet value.
 */
export class Truckload {
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ CONSTRUCTOR: constructor                                                   │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • key, truckNumber, and optional items.
   *
   * Returns
   *   • A new Truckload.
   *
   * Why / How
   *   A truck group needs both a safe grouping key and the display Truck #.
   *   It stores the key, Truck #, and its source-order item list.
   */
  constructor({ key, truckNumber, items = [] }) {
    this.key = key;
    this.truckNumber = truckNumber;
    this.items = items;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: addItem                                                            │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • item: InputItem.
   *
   * Returns
   *   • No returned value; updates the group.
   *
   * Why / How
   *   Truck groups are built while the reader walks source rows, so members
   *   must be appended consistently.
   *   It adds the item to the end of the member list.
   */
  addItem(item) {
    this.items.push(item);
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ GETTER: itemCount                                                          │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None.
   *
   * Returns
   *   • Number of items in this Truck #.
   *
   * Why / How
   *   Callers need a readable truck member count.
   *   It returns the item list length.
   */
  get itemCount() {
    return this.items.length;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: isFirstItem                                                        │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • item: InputItem.
   *
   * Returns
   *   • True only for the truck's Order item.
   *
   * Why / How
   *   The first source item for a Truck # is the FTL Order row.
   *   It compares the item with the first group member.
   */
  isFirstItem(item) {
    return this.items[0] === item;
  }

  /**
   * getCombinedItemGroups
   * Parameters: none. Returns: FtlItemGroup[] in first-appearance order.
   * Groups raw rows by case-insensitive Item Code within THIS truck, before
   * any Quantity, Total Weight, Freight Class, or Pallets calculations occur.
   * Missing/error Item Codes never combine with unrelated blank rows.
   */
  getCombinedItemGroups() {
    const byItemCode = new Map();
    const groups = [];

    this.items.forEach((item) => {
      const normalizedCode = normalizeText(item.itemCode);
      const isInvalid = !normalizedCode || normalizedCode === "error" ||
        normalizedCode.startsWith("error:");
      const key = isInvalid ? `missing-item-code-row-${item.sourceRow}` : normalizedCode;

      let group = byItemCode.get(key);
      if (!group) {
        group = new FtlItemGroup({ truck: this, firstItem: item });
        byItemCode.set(key, group);
        groups.push(group);
      } else {
        group.addItem(item);
      }
    });

    return groups;
  }

  /*
    In FTL, the Pallet source column identifies the truck's pallet total.
    The Order row gets the highest positive pallet value found for its Truck #.
    Commodity rows intentionally receive blank Pallets and Pallet Spaces.
  */
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getGreatestPalletCount                                             │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None; reads this.items.
   *
   * Returns
   *   • Greatest positive pallet count or ERROR when none is usable.
   *
   * Why / How
   *   The FTL Order row must show the truck's maximum Pallet number, even when
   *   that number appears later in the input.
   *   It converts every candidate pallet to a number, keeps positive values,
   *   and returns their maximum.
   */
  getGreatestPalletCount() {
    const usableCounts = this.items
      .map((item) => toFiniteNumber(item.pallet))
      .filter((value) => value !== null && value > 0);

    return usableCounts.length > 0
      ? Math.max(...usableCounts)
      : ERROR_VALUE;
  }
}
