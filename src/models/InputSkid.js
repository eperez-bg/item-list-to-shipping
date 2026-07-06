/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: InputSkid.js                                                         ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Represents one physical merged L/W/H group and calculates the LTL pallet   ║
 * ║ share for each item in that group.                                         ║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ LTL pallet fractions depend on physical merged L/W/H groups, which are     ║
 * ║ different from PO or Truck # groups.                                       ║
 * ║                                                                            ║
 * ║ MAIN CLASSES / FUNCTIONS                                                   ║
 * ║ • InputSkid — LTL pallet fractions depend on physical merged L/W/H         ║
 * ║ groups, which are different from PO or Truck # groups.                     ║
 * ║ • itemCount — Share calculations need the member count in one central      ║
 * ║ place.                                                                     ║
 * ║ • getPalletShareForItem — The template accepts at most three decimals,     ║
 * ║ but every physical group still needs to total exactly one pallet.          ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import { ERROR_VALUE } from "../utils/importIssues";

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ CLASS: InputSkid                                                           │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Constructor options for one physical dimension group.
 *
 * Returns
 *   • Instances used for LTL pallet and pallet-space calculations.
 *
 * Why / How
 *   LTL pallet fractions depend on physical merged L/W/H groups, which are
 *   different from PO or Truck # groups.
 *   The model holds the group rows, shared dimensions, and items, then
 *   computes an exact rounded share per item.
 */
export class InputSkid {
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ CONSTRUCTOR: constructor                                                   │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Source rows, dimensions, and items.
   *
   * Returns
   *   • A configured InputSkid instance.
   *
   * Why / How
   *   The physical grouping information should travel together instead of
   *   being recreated from raw Excel ranges later.
   *   It saves group boundaries, copied dimensions, and member items.
   */
  constructor({
    sourceStartRow,
    sourceEndRow,
    length,
    width,
    height,
    items,
  }) {
    this.sourceStartRow = sourceStartRow;
    this.sourceEndRow = sourceEndRow;
    this.length = length;
    this.width = width;
    this.height = height;
    this.items = items;
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
   *   • Number of items in the skid group.
   *
   * Why / How
   *   Share calculations need the member count in one central place.
   *   It returns the number of items in this physical group.
   */
  get itemCount() {
    return this.items.length;
  }

  /*
    Pallets and Pallet Spaces represent the physical L/W/H merged group.

    - A one-row, non-merged dimension group receives 1.000.
    - A merged dimension group shares one pallet across the actual item rows
      in that group.
    - Rounding is distributed so the group totals exactly 1.000 while no row
      has more than three decimal places.

    Examples:
      3 items -> 0.334, 0.333, 0.333
      6 items -> 0.167, 0.167, 0.167, 0.167, 0.166, 0.166
  */
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getPalletShareForItem                                              │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • item: member item; decimalPlaces: allowed precision.
   *
   * Returns
   *   • That item's pallet/pallet-space share, or ERROR when the item is not
   *   • in the group.
   *
   * Why / How
   *   The template accepts at most three decimals, but every physical group
   *   still needs to total exactly one pallet.
   *   It converts one pallet into integer units, distributes the remainder to
   *   early items, and converts back to decimals.
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
}
