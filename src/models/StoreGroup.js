/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: StoreGroup.js                                                     ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                          ║
 * ║ Represents one destination Store # and owns the generated LTL truck      ║
 * ║ groups for that store.                                                   ║
 * ║                                                                         ║
 * ║ WHY THIS FILE EXISTS                                                    ║
 * ║ LTL grouping now starts with Store #. A Store model keeps the grouping   ║
 * ║ rule easy to understand and gives future logic a clean place to split    ║
 * ║ one store into multiple trucks if management adds that requirement.      ║
 * ║                                                                         ║
 * ║ MAIN CLASSES / FUNCTIONS                                                ║
 * ║ • StoreGroup — Holds all generated trucks for one Store #.              ║
 * ║ • addItem — Adds a source item to the store's current truck.            ║
 * ║ • currentTruck — Returns the single active truck used by today's logic. ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import { StoreTruck } from "./StoreTruck";

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ CLASS: StoreGroup                                                        │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • key: normalized Store # key; storeNumber: original display value.
 *   • trucks: optional generated truck list.
 *
 * Returns
 *   • A StoreGroup that owns one or more StoreTruck objects.
 *
 * Why / How
 *   For now, one store creates one truck. The list shape is intentional so a
 *   later FTL-style split into multiple trucks per store can be added without
 *   changing the rest of the LTL output builder.
 */
export class StoreGroup {
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ CONSTRUCTOR: constructor                                                │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • key, storeNumber, and optional trucks.
   *
   * Returns
   *   • A StoreGroup instance.
   */
  constructor({ key, storeNumber, trucks = [] }) {
    this.key = key;
    this.storeNumber = storeNumber;
    this.trucks = trucks;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: addItem                                                         │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • item: InputItem for this Store #.
   *
   * Returns
   *   • No returned value; mutates the active StoreTruck.
   *
   * Why / How
   *   Today's rule has exactly one truck per store. This method sends every
   *   store item into that current truck in source order.
   */
  addItem(item) {
    this.currentTruck.addItem(item);
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ GETTER: currentTruck                                                    │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None.
   *
   * Returns
   *   • The single active StoreTruck for this store.
   *
   * Why / How
   *   Lazily creating the truck keeps StoreGroup construction light while still
   *   guaranteeing every store has a truck by the time the first item is added.
   */
  get currentTruck() {
    if (this.trucks.length === 0) {
      this.trucks.push(
        new StoreTruck({
          key: `${this.key}-truck-1`,
          storeKey: this.key,
          storeNumber: this.storeNumber,
        }),
      );
    }

    return this.trucks[0];
  }
}
