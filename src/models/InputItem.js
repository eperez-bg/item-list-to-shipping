import { ERROR_VALUE } from "../utils/importIssues";

export class InputItem {
  constructor({
    sourceRow,
    customerPo,
    targetStore,
    targetStoreSourceLocation,
    oldItemCode,
    quantity,
    pickupNumber,
    allocatedWeight = ERROR_VALUE,
    length = ERROR_VALUE,
    width = ERROR_VALUE,
    height = ERROR_VALUE,
    dimensionSourceStartRow = null,
    dimensionSourceEndRow = null,
  }) {
    this.sourceRow = sourceRow;
    this.customerPo = customerPo;
    this.targetStore = targetStore;
    this.targetStoreSourceLocation = targetStoreSourceLocation;
    this.oldItemCode = oldItemCode;
    this.quantity = quantity;
    this.pickupNumber = pickupNumber;

    /*
      This is the weight written to this one template row.

      - A non-merged G.W. cell belongs only to its own item.
      - A vertically merged G.W. cell is one shared total and is divided by
        the number of actual input item rows inside that merge range.
    */
    this.allocatedWeight = allocatedWeight;

    /*
      L/W/H may come from a vertically merged physical-dimension range. The
      complete dimensions are copied onto every input item in that range; they
      are never split between POs or items.
    */
    this.length = length;
    this.width = width;
    this.height = height;
    this.dimensionSourceStartRow = dimensionSourceStartRow;
    this.dimensionSourceEndRow = dimensionSourceEndRow;
  }

  assignDimensions({
    length,
    width,
    height,
    sourceStartRow,
    sourceEndRow,
  }) {
    this.length = length;
    this.width = width;
    this.height = height;
    this.dimensionSourceStartRow = sourceStartRow;
    this.dimensionSourceEndRow = sourceEndRow;
  }
}
