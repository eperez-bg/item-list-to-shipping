import { PurchaseOrder } from "./PurchaseOrder";
import { Truckload } from "./Truckload";
import { ERROR_VALUE } from "../utils/importIssues";
import { normalizeText } from "../utils/text";

export class InputShipment {
  constructor({ skids, issues = [] }) {
    this.skids = skids;
    this.issues = issues;

    this.items = this.skids.flatMap((skid) => skid.items);

    this.purchaseOrders = this.createPurchaseOrders();
    this.purchaseOrderByItem = this.createPurchaseOrderIndex();

    this.trucks = this.createTruckloads();
    this.truckByItem = this.createTruckIndex();
  }

  get itemCount() {
    return this.items.length;
  }

  get purchaseOrderCount() {
    return this.purchaseOrders.length;
  }

  get truckCount() {
    return this.trucks.length;
  }

  getPurchaseOrderForItem(item) {
    return this.purchaseOrderByItem.get(item) ?? null;
  }

  getTruckForItem(item) {
    return this.truckByItem.get(item) ?? null;
  }

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

  createPurchaseOrderIndex() {
    const index = new Map();

    this.purchaseOrders.forEach((purchaseOrder) => {
      purchaseOrder.items.forEach((item) => {
        index.set(item, purchaseOrder);
      });
    });

    return index;
  }

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

  createTruckIndex() {
    const index = new Map();

    this.trucks.forEach((truck) => {
      truck.items.forEach((item) => {
        index.set(item, truck);
      });
    });

    return index;
  }

  getPurchaseOrderKey(customerPo, sourceRow) {
    return this.getNormalizedGroupKey("missing-or-error-po", customerPo, sourceRow);
  }

  getTruckKey(truckNumber, sourceRow) {
    return this.getNormalizedGroupKey(
      "missing-or-error-truck",
      truckNumber,
      sourceRow,
    );
  }

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
