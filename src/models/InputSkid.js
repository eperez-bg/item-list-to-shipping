import { ERROR_VALUE } from "../utils/importIssues";
import { toFiniteNumber } from "../utils/text";

export class InputSkid {
  constructor({
    sourceStartRow,
    sourceEndRow,
    length,
    width,
    height,
    grossWeight,
    items,
  }) {
    this.sourceStartRow = sourceStartRow;
    this.sourceEndRow = sourceEndRow;
    this.length = length;
    this.width = width;
    this.height = height;
    this.grossWeight = grossWeight;
    this.items = items;
  }

  get itemCount() {
    return this.items.length;
  }

  get numericLength() {
    return positiveNumberOrNull(this.length);
  }

  get numericWidth() {
    return positiveNumberOrNull(this.width);
  }

  get numericHeight() {
    return positiveNumberOrNull(this.height);
  }

  get numericGrossWeight() {
    return positiveNumberOrNull(this.grossWeight);
  }

  get hasValidFreightMeasurements() {
    return Boolean(
      this.numericLength &&
      this.numericWidth &&
      this.numericHeight &&
      this.numericGrossWeight,
    );
  }

  get perItemWeight() {
    if (!this.numericGrossWeight || this.itemCount === 0) {
      return ERROR_VALUE;
    }

    return this.numericGrossWeight / this.itemCount;
  }

  get palletFraction() {
    return this.itemCount === 0 ? ERROR_VALUE : 1 / this.itemCount;
  }
}

function positiveNumberOrNull(value) {
  const numericValue = toFiniteNumber(value);
  return numericValue !== null && numericValue > 0 ? numericValue : null;
}
