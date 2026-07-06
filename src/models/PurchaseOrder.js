/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: PurchaseOrder.js                                                     ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Represents all items sharing one Customer PO so LTL can identify the       ║
 * ║ Order row and keep related Commodity rows together.                        ║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ LTL needs a PO-level object to decide which row is Order and to keep       ║
 * ║ every Commodity row beside it.                                             ║
 * ║                                                                            ║
 * ║ MAIN CLASSES / FUNCTIONS                                                   ║
 * ║ • PurchaseOrder — LTL needs a PO-level object to decide which row is       ║
 * ║ Order and to keep every Commodity row beside it.                           ║
 * ║ • addItem — The reader discovers items one at a time, so groups need one   ║
 * ║ controlled way to add members.                                             ║
 * ║ • itemCount — Consumers need a readable group size without accessing       ║
 * ║ implementation details.                                                    ║
 * ║ • getPalletShareForItem — This legacy helper remains available for         ║
 * ║ PO-based share experiments, even though current LTL uses InputSkid         ║
 * ║ physical shares.                                                           ║
 * ║ • isFirstItem — The first item in source order is the one that becomes     ║
 * ║ the LTL Order row for this PO.                                             ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import { ERROR_VALUE } from "../utils/importIssues";

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ CLASS: PurchaseOrder                                                       │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Constructor options for a normalized PO group.
 *
 * Returns
 *   • Instances used by LTL output and preview generation.
 *
 * Why / How
 *   LTL needs a PO-level object to decide which row is Order and to keep
 *   every Commodity row beside it.
 *   The model stores a source-order list of member items and exposes simple
 *   group queries.
 */
export class PurchaseOrder {
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ CONSTRUCTOR: constructor                                                   │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • key, customerPo, and optional starting items.
   *
   * Returns
   *   • A new PurchaseOrder.
   *
   * Why / How
   *   A PO group needs both an internal safe key and the original display PO
   *   value.
   *   It stores the key, display value, and member items.
   */
  constructor({ key, customerPo, items = [] }) {
    this.key = key;
    this.customerPo = customerPo;
    this.items = items;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: addItem                                                            │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • item: InputItem to add.
   *
   * Returns
   *   • No returned value; updates the group.
   *
   * Why / How
   *   The reader discovers items one at a time, so groups need one controlled
   *   way to add members.
   *   It appends the item, preserving source order for Order/Commodity
   *   decisions.
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
   *   • Number of items in this PO.
   *
   * Why / How
   *   Consumers need a readable group size without accessing implementation
   *   details.
   *   It returns the member array length.
   */
  get itemCount() {
    return this.items.length;
  }

  /*
    Returns this item's share of one pallet / one pallet space for its PO.

    The share uses at most three decimal places and distributes the rounding
    remainder across the first rows. That means every PO totals exactly 1.000:

      3 items -> 0.334, 0.333, 0.333
      6 items -> 0.167, 0.167, 0.167, 0.167, 0.166, 0.166
  */
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getPalletShareForItem                                              │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • item and decimalPlaces.
   *
   * Returns
   *   • Rounded PO share or ERROR for a non-member.
   *
   * Why / How
   *   This legacy helper remains available for PO-based share experiments,
   *   even though current LTL uses InputSkid physical shares.
   *   It applies the same integer-unit rounding strategy used by skid shares.
   */
  getPalletShareForItem(item, decimalPlaces = 3) {
    const itemIndex = this.items.indexOf(item);

    if (itemIndex < 0 || this.itemCount === 0) {
      return ERROR_VALUE;
    }

    const scale = 10 ** decimalPlaces;
    const baseUnits = Math.floor(scale / this.itemCount);
    const remainderUnits = scale - baseUnits * this.itemCount;
    const shareUnits = baseUnits + (itemIndex < remainderUnits ? 1 : 0);

    return shareUnits / scale;
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
   *   • True only for the group's first item.
   *
   * Why / How
   *   The first item in source order is the one that becomes the LTL Order row
   *   for this PO.
   *   It compares the item with the first stored member.
   */
  isFirstItem(item) {
    return this.items[0] === item;
  }
}
