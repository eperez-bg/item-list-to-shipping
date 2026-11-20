/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: StoreTruck.js                                                    ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                          ║
 * ║ Represents one LTL output truck/group for a destination Store # and     ║
 * ║ creates the combined item groups that will become template rows.        ║
 * ║                                                                         ║
 * ║ WHY THIS FILE EXISTS                                                    ║
 * ║ The current LTL business rule says one Store # becomes one generated     ║
 * ║ truck/group for now. Inside that truck, duplicate item rows with the    ║
 * ║ same Item Code and same CustomerPO should be combined before output.    ║
 * ║ Keeping this in the truck model keeps grouping rules out of the Excel   ║
 * ║ writer and gives future multi-truck-per-store logic a clean home.       ║
 * ║                                                                         ║
 * ║ MAIN CLASSES / FUNCTIONS                                                ║
 * ║ • StoreTruck — Holds source-order items for one generated LTL truck.    ║
 * ║ • StoreTruckItemGroup — Represents one combined output item row.        ║
 * ║ • addItem — Appends an input item while preserving source order.        ║
 * ║ • getCombinedItemGroups — Combines same Item Code + CustomerPO items.   ║
 * ║ • getItemCombinationKey — Builds the duplicate-item grouping key.       ║
 * ║ • orderItem — Returns the first item, which supplies group-level values.║
 * ║ • isFirstItem — Identifies the original Order item for this truck.      ║
 * ║ • targetPo / pickupNumber — Read group-level values from the Order row. ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import { ERROR_VALUE } from "../utils/importIssues";
import { isBlank, normalizeText } from "../utils/text";

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ CLASS: StoreTruck                                                       │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • key: stable group key; storeKey/storeNumber: parent store identity.
 *   • items: optional starting source items.
 *
 * Returns
 *   • A generated LTL truck/group model.
 *
 * Why / How
 *   LTL rows are grouped by Store #. The first item found for a Store # is the
 *   Order item; every later item for that same store belongs under it. Before
 *   the truck is written to the template, duplicate items with the same Item
 *   Code and same CustomerPO are combined into one output row.
 */
export class StoreTruck {
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ CONSTRUCTOR: constructor                                                │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • key, storeKey, storeNumber, and optional source items.
   *
   * Returns
   *   • A StoreTruck instance.
   *
   * Why / How
   *   Truck state should be initialized consistently even if future code passes
   *   prebuilt items. Each item is routed through addItem to preserve behavior.
   */
  constructor({ key, storeKey, storeNumber, items = [] }) {
    this.key = key;
    this.storeKey = storeKey;
    this.storeNumber = storeNumber;
    this.items = [];

    items.forEach((item) => this.addItem(item));
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: addItem                                                         │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • item: InputItem belonging to this Store #.
   *
   * Returns
   *   • No returned value; updates this.items.
   *
   * Why / How
   *   Source order decides which raw row is the Order item. This method appends
   *   items so the first source item remains the Order item even if duplicates
   *   are later combined into that Order output row.
   */
  addItem(item) {
    this.items.push(item);
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
   *   • Number of raw input items in this generated truck.
   */
  get itemCount() {
    return this.items.length;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ GETTER: orderItem                                                       │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None.
   *
   * Returns
   *   • First raw item in the truck, or null when empty.
   *
   * Why / How
   *   Group-level fields like the first Target PO# columns and Pickup Number
   *   must come from the original Order row, even when that Order row is later
   *   combined with matching Commodity rows.
   */
  get orderItem() {
    return this.items[0] ?? null;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ GETTER: commodityItems                                                  │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None.
   *
   * Returns
   *   • All raw input items after the original Order item.
   */
  get commodityItems() {
    return this.items.slice(1);
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: isFirstItem                                                     │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • item: InputItem candidate.
   *
   * Returns
   *   • True when item is the raw Order item for this truck.
   */
  isFirstItem(item) {
    return this.orderItem === item;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: isOrderGroup                                                    │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • group: StoreTruckItemGroup created by getCombinedItemGroups.
   *
   * Returns
   *   • True when the combined group contains the raw Order item.
   *
   * Why / How
   *   If the raw Order item combines with matching Commodity items, the single
   *   combined output row must still be the truck's Order row.
   */
  isOrderGroup(group) {
    return group.includesItem(this.orderItem);
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getCombinedItemGroups                                           │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None; reads this.items.
   *
   * Returns
   *   • StoreTruckItemGroup[] in first-appearance order.
   *
   * Why / How
   *   Management wants repeated items on the same Store/Truck to collapse into
   *   one template row only when both Item Code and CustomerPO match. The method
   *   walks the raw truck items once, creates a grouping key for each item, and
   *   preserves the first source row as the representative row for that group.
   */
  getCombinedItemGroups() {
    const groupsByKey = new Map();

    this.items.forEach((item) => {
      const key = this.getItemCombinationKey(item);

      if (!groupsByKey.has(key)) {
        groupsByKey.set(
          key,
          new StoreTruckItemGroup({
            key,
            representativeItem: item,
          }),
        );
      }

      groupsByKey.get(key).addItem(item);
    });

    return Array.from(groupsByKey.values());
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getItemCombinationKey                                           │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • item: InputItem being placed into a combined item group.
   *
   * Returns
   *   • A normalized key, or a row-specific key for unsafe missing values.
   *
   * Why / How
   *   Combining must be based on the input Item Code and CustomerPO. Missing or
   *   ERROR values are isolated by source row so unrelated bad rows do not get
   *   collapsed together by accident.
   */
  getItemCombinationKey(item) {
    const itemCodeKey = this.normalizeCombinationPart(item.itemCode);
    const customerPoKey = this.normalizeCombinationPart(item.customerPo);

    if (!itemCodeKey || !customerPoKey) {
      return `uncombined-row-${item.sourceRow}`;
    }

    return `item-${itemCodeKey}__po-${customerPoKey}`;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: normalizeCombinationPart                                        │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • value: raw Item Code or CustomerPO value.
   *
   * Returns
   *   • Normalized text safe for grouping, or blank when grouping is unsafe.
   *
   * Why / How
   *   This keeps the combine check case-insensitive and spacing-insensitive,
   *   but still prevents blank/ERROR values from accidentally matching.
   */
  normalizeCombinationPart(value) {
    const normalizedValue = normalizeText(value);

    if (!normalizedValue || normalizedValue === normalizeText(ERROR_VALUE)) {
      return "";
    }

    return normalizedValue;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ GETTER: targetPo                                                        │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None.
   *
   * Returns
   *   • CustomerPO from the raw Order item, or blank.
   *
   * Why / How
   *   The first Target PO# columns use the Order item's PO for every item in the
   *   same Store/Truck group, even when Commodity rows have different POs.
   */
  get targetPo() {
    return this.orderItem?.customerPo ?? "";
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ GETTER: pickupNumber                                                    │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None.
   *
   * Returns
   *   • IF value from the raw Order item, or blank.
   *
   * Why / How
   *   Pickup Number should be identical for the whole truck group and should
   *   come from the original Order item.
   */
  get pickupNumber() {
    return this.orderItem?.pickupNumber ?? "";
  }
}

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ CLASS: StoreTruckItemGroup                                              │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • key: duplicate-item key; representativeItem: first item in the group.
 *
 * Returns
 *   • A combined item group that will become one LTL output row.
 *
 * Why / How
 *   One output row may represent multiple raw source rows. This small model
 *   keeps those raw rows together while exposing the first row as the source for
 *   item-level values that are not mathematically combined, such as IF number.
 */
export class StoreTruckItemGroup {
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ CONSTRUCTOR: constructor                                                │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • key and representativeItem.
   *
   * Returns
   *   • A StoreTruckItemGroup instance.
   */
  constructor({ key, representativeItem }) {
    this.key = key;
    this.representativeItem = representativeItem;
    this.items = [];
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: addItem                                                         │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • item: InputItem with matching Item Code and CustomerPO.
   *
   * Returns
   *   • No returned value; appends to this.items.
   */
  addItem(item) {
    this.items.push(item);
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: includesItem                                                    │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • item: InputItem candidate.
   *
   * Returns
   *   • True when the raw item is part of this combined group.
   */
  includesItem(item) {
    return Boolean(item) && this.items.includes(item);
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ GETTER: sourceRows                                                      │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None.
   *
   * Returns
   *   • Comma-separated source row list for logs and preview display.
   */
  get sourceRows() {
    return this.items.map((item) => item.sourceRow).join(", ");
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
   *   • Number of raw source rows represented by this output group.
   */
  get itemCount() {
    return this.items.length;
  }
}
