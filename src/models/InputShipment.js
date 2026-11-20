/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: InputShipment.js                                                  ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                          ║
 * ║ Owns the parsed collection of input items and builds the grouping models║
 * ║ used by the LTL and FTL output flows.                                    ║
 * ║                                                                         ║
 * ║ WHY THIS FILE EXISTS                                                    ║
 * ║ The workbook reader should only normalize rows and physical merged-cell ║
 * ║ groups. Business groupings, such as LTL Store/Truck groups and FTL      ║
 * ║ Truck # groups, belong in one shared shipment model.                    ║
 * ║                                                                         ║
 * ║ MAIN CLASSES / FUNCTIONS                                                ║
 * ║ • InputShipment — Parsed shipment plus grouping helpers.                ║
 * ║ • createStores — Groups LTL rows by normalized Store #.                 ║
 * ║ • createStoreTrucks — Flattens each StoreGroup truck list for LTL.      ║
 * ║ • createPurchaseOrders — Kept for compatibility/diagnostics.            ║
 * ║ • createTruckloads — Groups FTL rows by Truck #.                        ║
 * ║ • get*Key — Builds safe keys while isolating missing/error values.      ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import { PurchaseOrder } from "./PurchaseOrder";
import { StoreGroup } from "./StoreGroup";
import { Truckload } from "./Truckload";
import { ERROR_VALUE } from "../utils/importIssues";
import { isBlank, normalizeText } from "../utils/text";

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ CLASS: InputShipment                                                    │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • skids: physical merged-dimension groups; issues: recoverable import log.
 *
 * Returns
 *   • An InputShipment instance with ready-to-use LTL and FTL groupings.
 *
 * Why / How
 *   Output builders should not repeatedly scan the raw row list to determine
 *   groups. The constructor flattens all items once, then eagerly creates Store
 *   groups for LTL, Truck # groups for FTL, and lookup indexes for previews.
 */
export class InputShipment {
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ CONSTRUCTOR: constructor                                                │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • skids and optional issues from InputWorkbookReader.
   *
   * Returns
   *   • A populated InputShipment instance.
   *
   * Why / How
   *   The shipment should be ready immediately after parsing. This constructor
   *   saves the physical groups and builds business group collections in source
   *   first-appearance order.
   */
  constructor({ skids, issues = [] }) {
    this.skids = skids;
    this.issues = issues;
    this.items = this.skids.flatMap((skid) => skid.items);

    this.purchaseOrders = this.createPurchaseOrders();
    this.purchaseOrderByItem = this.createPurchaseOrderIndex();

    this.stores = this.createStores();
    this.storeByItem = this.createStoreIndex();
    this.storeTrucks = this.createStoreTrucks();
    this.storeTruckByItem = this.createStoreTruckIndex();

    this.trucks = this.createTruckloads();
    this.truckByItem = this.createTruckIndex();
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ GETTER: itemCount                                                       │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None.
   *
   * Returns
   *   • Number of parsed input items.
   */
  get itemCount() {
    return this.items.length;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ GETTER: purchaseOrderCount                                              │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None.
   *
   * Returns
   *   • Number of CustomerPO groups.
   */
  get purchaseOrderCount() {
    return this.purchaseOrders.length;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ GETTER: storeCount                                                      │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None.
   *
   * Returns
   *   • Number of normalized Store # groups for LTL.
   */
  get storeCount() {
    return this.stores.length;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ GETTER: storeTruckCount                                                 │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None.
   *
   * Returns
   *   • Number of generated StoreTruck groups for LTL.
   */
  get storeTruckCount() {
    return this.storeTrucks.length;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ GETTER: truckCount                                                      │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None.
   *
   * Returns
   *   • Number of Truck # groups for FTL.
   */
  get truckCount() {
    return this.trucks.length;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getPurchaseOrderForItem                                         │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • item: InputItem.
   *
   * Returns
   *   • PurchaseOrder or null.
   */
  getPurchaseOrderForItem(item) {
    return this.purchaseOrderByItem.get(item) ?? null;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getStoreForItem                                                 │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • item: InputItem.
   *
   * Returns
   *   • StoreGroup or null.
   */
  getStoreForItem(item) {
    return this.storeByItem.get(item) ?? null;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getStoreTruckForItem                                            │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • item: InputItem.
   *
   * Returns
   *   • StoreTruck or null.
   */
  getStoreTruckForItem(item) {
    return this.storeTruckByItem.get(item) ?? null;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getTruckForItem                                                 │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • item: InputItem.
   *
   * Returns
   *   • Truckload or null.
   */
  getTruckForItem(item) {
    return this.truckByItem.get(item) ?? null;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: createPurchaseOrders                                            │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None; reads this.items.
   *
   * Returns
   *   • PurchaseOrder instances in first-appearance order.
   *
   * Why / How
   *   PO groups are no longer the LTL Order/Commodity grouping, but they remain
   *   useful for diagnostics and any future PO-based reports.
   */
  createPurchaseOrders() {
    const groups = new Map();

    this.items.forEach((item) => {
      const key = this.getPurchaseOrderKey(item.customerPo, item.sourceRow);

      if (!groups.has(key)) {
        groups.set(key, new PurchaseOrder({ key, customerPo: item.customerPo }));
      }

      groups.get(key).addItem(item);
    });

    return Array.from(groups.values());
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: createPurchaseOrderIndex                                        │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None; reads this.purchaseOrders.
   *
   * Returns
   *   • Map<InputItem, PurchaseOrder>.
   */
  createPurchaseOrderIndex() {
    const index = new Map();

    this.purchaseOrders.forEach((purchaseOrder) => {
      purchaseOrder.items.forEach((item) => index.set(item, purchaseOrder));
    });

    return index;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: createStores                                                    │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None; reads this.items.
   *
   * Returns
   *   • StoreGroup instances in first-appearance order.
   *
   * Why / How
   *   LTL now groups by destination Store #. The first item for a Store # starts
   *   that store's generated truck, and later rows for that store are appended
   *   as Commodities beneath the same Order row.
   */
  createStores() {
    const groups = new Map();

    this.items.forEach((item) => {
      const key = this.getStoreKey(item.targetStore, item.sourceRow);

      if (!groups.has(key)) {
        groups.set(
          key,
          new StoreGroup({
            key,
            storeNumber: item.targetStore,
          }),
        );
      }

      groups.get(key).addItem(item);
    });

    return Array.from(groups.values());
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: createStoreIndex                                                │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None; reads this.stores.
   *
   * Returns
   *   • Map<InputItem, StoreGroup>.
   */
  createStoreIndex() {
    const index = new Map();

    this.stores.forEach((store) => {
      store.trucks.forEach((truck) => {
        truck.items.forEach((item) => index.set(item, store));
      });
    });

    return index;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: createStoreTrucks                                               │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None; reads this.stores.
   *
   * Returns
   *   • Flattened StoreTruck instances in store first-appearance order.
   */
  createStoreTrucks() {
    return this.stores.flatMap((store) => store.trucks);
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: createStoreTruckIndex                                           │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None; reads this.storeTrucks.
   *
   * Returns
   *   • Map<InputItem, StoreTruck>.
   */
  createStoreTruckIndex() {
    const index = new Map();

    this.storeTrucks.forEach((truck) => {
      truck.items.forEach((item) => index.set(item, truck));
    });

    return index;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: createTruckloads                                                │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None; reads this.items.
   *
   * Returns
   *   • Truckload instances in first-appearance Truck # order.
   *
   * Why / How
   *   FTL is intentionally restored to its Truck # grouping. This keeps the FTL
   *   flow separate from the new LTL Store # grouping.
   */
  createTruckloads() {
    const groups = new Map();

    this.items.forEach((item) => {
      const key = this.getTruckKey(item.truckNumber, item.sourceRow);

      if (!groups.has(key)) {
        groups.set(key, new Truckload({ key, truckNumber: item.truckNumber }));
      }

      groups.get(key).addItem(item);
    });

    return Array.from(groups.values());
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: createTruckIndex                                                │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None; reads this.trucks.
   *
   * Returns
   *   • Map<InputItem, Truckload>.
   */
  createTruckIndex() {
    const index = new Map();

    this.trucks.forEach((truck) => {
      truck.items.forEach((item) => index.set(item, truck));
    });

    return index;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getPurchaseOrderKey                                             │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • customerPo and the row where it was found.
   *
   * Returns
   *   • A safe PO grouping key.
   */
  getPurchaseOrderKey(customerPo, sourceRow) {
    return this.getNormalizedGroupKey("missing-or-error-po", customerPo, sourceRow);
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getStoreKey                                                     │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • storeNumber and the row where it was found.
   *
   * Returns
   *   • A safe Store # grouping key.
   *
   * Why / How
   *   Store values can arrive as 1310, 01310, 1310.0, or text containing a
   *   single store number. The key extracts that one number and strips leading
   *   zeros so equivalent store values group together.
   */
  getStoreKey(storeNumber, sourceRow) {
    const normalizedStore = this.getNormalizedStoreNumber(storeNumber);

    if (!normalizedStore || normalizedStore === normalizeText(ERROR_VALUE)) {
      return `missing-or-error-store-row-${sourceRow}`;
    }

    return `store-${normalizedStore}`;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getTruckKey                                                     │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • truckNumber and the row where it was found.
   *
   * Returns
   *   • A safe Truck # grouping key for FTL.
   */
  getTruckKey(truckNumber, sourceRow) {
    return this.getNormalizedGroupKey("missing-or-error-truck", truckNumber, sourceRow);
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getNormalizedStoreNumber                                        │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • value: raw Store # cell value.
   *
   * Returns
   *   • A normalized digit string with leading zeros removed, or blank.
   *
   * Why / How
   *   Store grouping should not split 1310, 01310, and 1310.0 into different
   *   trucks. It accepts direct numeric-looking values or one number embedded
   *   in a label, but refuses values with multiple unrelated numbers.
   */
  getNormalizedStoreNumber(value) {
    if (isBlank(value)) return "";

    if (typeof value === "number") {
      if (!Number.isSafeInteger(value) || value < 0) return "";
      return String(value);
    }

    const text = String(value)
      .normalize("NFKC")
      .replace(/\u00a0/g, " ")
      .replace(/,/g, "")
      .trim();

    const directMatch = text.match(/^0*(\d+)(?:\.0+)?$/);
    if (directMatch) return directMatch[1] || "0";

    const numberGroups = text.match(/\d+/g) ?? [];
    if (numberGroups.length !== 1) return "";

    return numberGroups[0].replace(/^0+/, "") || "0";
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getNormalizedGroupKey                                           │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • missingPrefix, raw group value, and sourceRow.
   *
   * Returns
   *   • A normalized, safe Map key.
   *
   * Why / How
   *   Grouping needs case-insensitive matching but must isolate missing/error
   *   values so unrelated bad rows do not become one group by accident.
   */
  getNormalizedGroupKey(missingPrefix, value, sourceRow) {
    const normalizedValue = normalizeText(value);

    if (!normalizedValue || normalizedValue === normalizeText(ERROR_VALUE)) {
      return `${missingPrefix}-row-${sourceRow}`;
    }

    return normalizedValue;
  }
}
