import { ERROR_VALUE } from "../utils/importIssues";

export class PurchaseOrder {
  constructor({ key, customerPo, items = [] }) {
    this.key = key;
    this.customerPo = customerPo;
    this.items = items;
  }

  addItem(item) {
    this.items.push(item);
  }

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

  isFirstItem(item) {
    return this.items[0] === item;
  }
}
