/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: InputShipment.js                                                     ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Owns the parsed collection of input items and builds the PO and truck      ║
 * ║ groupings used by LTL and FTL output.                                      ║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ Both workflows need the same parsed items plus different ways to group     ║
 * ║ them, so grouping belongs in a shared shipment model.                      ║
 * ║                                                                            ║
 * ║ MAIN CLASSES / FUNCTIONS                                                   ║
 * ║ • InputShipment — Both workflows need the same parsed items plus           ║
 * ║ different ways to group them, so grouping belongs in a shared shipment     ║
 * ║ model.                                                                     ║
 * ║ • itemCount — The UI needs a reliable summary count without duplicating    ║
 * ║ the flattening logic.                                                      ║
 * ║ • purchaseOrderCount — LTL status messaging needs the number of logical    ║
 * ║ PO groups.                                                                 ║
 * ║ • truckCount — FTL status messaging needs the number of Truck # groups.    ║
 * ║ • getPurchaseOrderForItem — Callers should not have to scan every PO       ║
 * ║ group just to find an item's group.                                        ║
 * ║ • getTruckForItem — Callers should not have to scan every Truck # group    ║
 * ║ just to find an item's group.                                              ║
 * ║ • createPurchaseOrders — LTL output must group all same-PO items           ║
 * ║ together even when the input sheet has them in separate physical groups.   ║
 * ║ • createPurchaseOrderIndex — Fast lookup avoids repeated nested loops      ║
 * ║ during preview or output creation.                                         ║
 * ║ • createTruckloads — FTL output must group all same-truck items together   ║
 * ║ even when they are separated in the input sheet.                           ║
 * ║ • createTruckIndex — Fast lookup avoids repeated scans of truck groups.    ║
 * ║ • getPurchaseOrderKey — Missing or invalid PO values must not make         ║
 * ║ unrelated rows look like one order.                                        ║
 * ║ • getTruckKey — Missing or invalid Truck # values must not combine         ║
 * ║ unrelated FTL rows.                                                        ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import { PurchaseOrder } from "./PurchaseOrder";
import { Truckload } from "./Truckload";
import { ERROR_VALUE } from "../utils/importIssues";
import { normalizeText } from "../utils/text";

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ CLASS: InputShipment                                                       │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • skids: physical groups; issues: recoverable import issues.
 *
 * Returns
 *   • An InputShipment instance with grouping helpers.
 *
 * Why / How
 *   Both workflows need the same parsed items plus different ways to group
 *   them, so grouping belongs in a shared shipment model.
 *   It flattens skid items once, then builds PO and Truck # collections plus
 *   fast item-to-group indexes.
 */
export class InputShipment {
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ CONSTRUCTOR: constructor                                                   │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • skids and optional issues from InputWorkbookReader.
   *
   * Returns
   *   • A populated InputShipment instance.
   *
   * Why / How
   *   The shipment should be fully ready for either LTL or FTL immediately
   *   after construction.
   *   It saves the raw collections and eagerly builds all grouping and lookup
   *   structures.
   */
  constructor({ skids, issues = [] }) {
    this.skids = skids;
    this.issues = issues;

    this.items = this.skids.flatMap((skid) => skid.items);

    this.purchaseOrders = this.createPurchaseOrders();
    this.purchaseOrderByItem = this.createPurchaseOrderIndex();

    this.trucks = this.createTruckloads();
    this.truckByItem = this.createTruckIndex();
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
   *   • Number of input items.
   *
   * Why / How
   *   The UI needs a reliable summary count without duplicating the flattening
   *   logic.
   *   It returns the length of the normalized items collection.
   */
  get itemCount() {
    return this.items.length;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ GETTER: purchaseOrderCount                                                 │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None.
   *
   * Returns
   *   • Number of PO groups.
   *
   * Why / How
   *   LTL status messaging needs the number of logical PO groups.
   *   It returns the current purchaseOrders array length.
   */
  get purchaseOrderCount() {
    return this.purchaseOrders.length;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ GETTER: truckCount                                                         │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None.
   *
   * Returns
   *   • Number of truck groups.
   *
   * Why / How
   *   FTL status messaging needs the number of Truck # groups.
   *   It returns the current trucks array length.
   */
  get truckCount() {
    return this.trucks.length;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getPurchaseOrderForItem                                            │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • item: an InputItem.
   *
   * Returns
   *   • PurchaseOrder or null.
   *
   * Why / How
   *   Callers should not have to scan every PO group just to find an item's
   *   group.
   *   It performs a Map lookup created during construction.
   */
  getPurchaseOrderForItem(item) {
    return this.purchaseOrderByItem.get(item) ?? null;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getTruckForItem                                                    │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • item: an InputItem.
   *
   * Returns
   *   • Truckload or null.
   *
   * Why / How
   *   Callers should not have to scan every Truck # group just to find an
   *   item's group.
   *   It performs a Map lookup created during construction.
   */
  getTruckForItem(item) {
    return this.truckByItem.get(item) ?? null;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: createPurchaseOrders                                               │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None; reads this.items.
   *
   * Returns
   *   • PurchaseOrder instances in first-appearance order.
   *
   * Why / How
   *   LTL output must group all same-PO items together even when the input
   *   sheet has them in separate physical groups.
   *   It creates one PurchaseOrder per normalized PO key and appends items in
   *   source order.
   */
  createPurchaseOrders() {
    const groups = new Map();

    this.items.forEach((item) => {
      const key = this.getPurchaseOrderKey(item.customerPo, item.sourceRow);

      if (!groups.has(key)) {
        groups.set(
          key,
          new PurchaseOrder({
            key,
            customerPo: item.customerPo,
          }),
        );
      }

      groups.get(key).addItem(item);
    });

    return Array.from(groups.values());
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: createPurchaseOrderIndex                                           │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None; reads this.purchaseOrders.
   *
   * Returns
   *   • Map<InputItem, PurchaseOrder>.
   *
   * Why / How
   *   Fast lookup avoids repeated nested loops during preview or output
   *   creation.
   *   It creates a Map from each InputItem to its containing PurchaseOrder.
   */
  createPurchaseOrderIndex() {
    const index = new Map();

    this.purchaseOrders.forEach((purchaseOrder) => {
      purchaseOrder.items.forEach((item) => {
        index.set(item, purchaseOrder);
      });
    });

    return index;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: createTruckloads                                                   │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None; reads this.items.
   *
   * Returns
   *   • Truckload instances in first-appearance order.
   *
   * Why / How
   *   FTL output must group all same-truck items together even when they are
   *   separated in the input sheet.
   *   It creates one Truckload per normalized Truck # key and appends items in
   *   source order.
   */
  createTruckloads() {
    const groups = new Map();

    this.items.forEach((item) => {
      const key = this.getTruckKey(item.truckNumber, item.sourceRow);

      if (!groups.has(key)) {
        groups.set(
          key,
          new Truckload({
            key,
            truckNumber: item.truckNumber,
          }),
        );
      }

      groups.get(key).addItem(item);
    });

    return Array.from(groups.values());
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: createTruckIndex                                                   │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None; reads this.trucks.
   *
   * Returns
   *   • Map<InputItem, Truckload>.
   *
   * Why / How
   *   Fast lookup avoids repeated scans of truck groups.
   *   It creates a Map from each InputItem to its containing Truckload.
   */
  createTruckIndex() {
    const index = new Map();

    this.trucks.forEach((truck) => {
      truck.items.forEach((item) => {
        index.set(item, truck);
      });
    });

    return index;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getPurchaseOrderKey                                                │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • customerPo and the row where it was found.
   *
   * Returns
   *   • A safe PO grouping key.
   *
   * Why / How
   *   Missing or invalid PO values must not make unrelated rows look like one
   *   order.
   *   It delegates to the normalized-key helper with a PO-specific missing
   *   prefix.
   */
  getPurchaseOrderKey(customerPo, sourceRow) {
    return this.getNormalizedGroupKey("missing-or-error-po", customerPo, sourceRow);
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getTruckKey                                                        │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • truckNumber and the row where it was found.
   *
   * Returns
   *   • A safe truck grouping key.
   *
   * Why / How
   *   Missing or invalid Truck # values must not combine unrelated FTL rows.
   *   It delegates to the normalized-key helper with a truck-specific missing
   *   prefix.
   */
  getTruckKey(truckNumber, sourceRow) {
    return this.getNormalizedGroupKey(
      "missing-or-error-truck",
      truckNumber,
      sourceRow,
    );
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getNormalizedGroupKey                                              │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • missingPrefix, raw group value, and sourceRow.
   *
   * Returns
   *   • A normalized, safe Map key.
   *
   * Why / How
   *   Grouping needs case-insensitive matching but must preserve row isolation
   *   for missing/error values.
   *   It normalizes usable text; otherwise it builds a unique key using the
   *   row number.
   */
  getNormalizedGroupKey(missingPrefix, value, sourceRow) {
    const normalizedValue = normalizeText(value);

    // Unknown values must never get grouped together. Each bad row is its own
    // Order row instead of accidentally becoming Commodity.
    if (!normalizedValue || normalizedValue === normalizeText(ERROR_VALUE)) {
      return `${missingPrefix}-row-${sourceRow}`;
    }

    return normalizedValue;
  }
}
