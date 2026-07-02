import { ERROR_VALUE } from "../utils/importIssues";

export class InputItem {
  constructor({
    sourceRow,
    customerPo,
    targetStore,
    targetStoreSourceLocation,
    oldItemCode,
    quantity,
    pickupNumber = "",
    truckNumber = "",
    pallet = "",
    chinaPo = "",
    container = "",
    isAdditionalPallet = false,
    grossWeight = ERROR_VALUE,
    grossWeightSourceLocation = "",
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

    // FTL-only source values. They remain empty in LTL mode.
    this.truckNumber = truckNumber;
    this.pallet = pallet;
    this.chinaPo = chinaPo;
    this.container = container;
    this.isAdditionalPallet = isAdditionalPallet;

    // The full source G.W. value seen by this item. FTL uses it on the Truck's
    // Order row. It is not split.
    this.grossWeight = grossWeight;
    this.grossWeightSourceLocation = grossWeightSourceLocation;

    // LTL freight classification keeps the prior behavior: a vertically
    // merged G.W. is split across only the actual items in that G.W. range.
    this.allocatedWeight = allocatedWeight;

    // L/W/H can come from a vertically merged physical-dimension range. The
    // full dimensions are copied to every item in that physical group.
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
