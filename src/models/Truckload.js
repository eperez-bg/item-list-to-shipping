import { ERROR_VALUE } from "../utils/importIssues";
import { toFiniteNumber } from "../utils/text";

export class Truckload {
  constructor({ key, truckNumber, items = [] }) {
    this.key = key;
    this.truckNumber = truckNumber;
    this.items = items;
  }

  addItem(item) {
    this.items.push(item);
  }

  get itemCount() {
    return this.items.length;
  }

  isFirstItem(item) {
    return this.items[0] === item;
  }

  /*
    In FTL, the Pallet source column identifies the truck's pallet total.
    The Order row gets the highest positive pallet value found for its Truck #.
    Commodity rows intentionally receive blank Pallets and Pallet Spaces.
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
