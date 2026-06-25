import { PurchaseOrder } from "./PurchaseOrder";
import { ERROR_VALUE } from "../utils/importIssues";
import { normalizeText } from "../utils/text";

export class InputShipment {
  constructor({ skids, issues = [] }) {
    this.skids = skids;
    this.issues = issues;

    this.items = this.skids.flatMap((skid) => skid.items);
    this.purchaseOrders = this.createPurchaseOrders();
    this.purchaseOrderByItem = this.createPurchaseOrderIndex();
  }

  get itemCount() {
    return this.items.length;
  }

  get purchaseOrderCount() {
    return this.purchaseOrders.length;
  }

  getPurchaseOrderForItem(item) {
    return this.purchaseOrderByItem.get(item) ?? null;
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

  getPurchaseOrderKey(customerPo, sourceRow) {
    const normalizedPo = normalizeText(customerPo);

    /*
      Missing / ERROR POs must never be combined. Each row is treated as its
      own PO group so separate unknown orders do not accidentally become
      Commodity rows or share pallet fractions.
    */
    if (!normalizedPo || normalizedPo === normalizeText(ERROR_VALUE)) {
      return `missing-or-error-po-row-${sourceRow}`;
    }

    return normalizedPo;
  }
}
