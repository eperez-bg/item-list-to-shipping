export class InputItem {
  constructor({
    sourceRow,
    customerPo,
    targetStore,
    oldItemCode,
    quantity,
    pickupNumber,
  }) {
    this.sourceRow = sourceRow;
    this.customerPo = customerPo;
    this.targetStore = targetStore;
    this.oldItemCode = oldItemCode;
    this.quantity = quantity;
    this.pickupNumber = pickupNumber;
  }
}
