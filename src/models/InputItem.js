/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: InputItem.js                                                      ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                          ║
 * ║ Defines the normalized in-memory record for one input row after it has  ║
 * ║ been read from Excel.                                                   ║
 * ║                                                                         ║
 * ║ WHY THIS FILE EXISTS                                                    ║
 * ║ A single source row has many related values, and a model keeps          ║
 * ║ downstream services from passing loose, error-prone objects.            ║
 * ║                                                                         ║
 * ║ MAIN CLASSES / FUNCTIONS                                                ║
 * ║ • InputItem — Stores one normalized source row plus calculated values.  ║
 * ║ • assignDimensions — Copies merged L/W/H values onto every item in the  ║
 * ║   physical dimension group.                                             ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import { ERROR_VALUE } from "../utils/importIssues";

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ CLASS: InputItem                                                        │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Constructor options describing one source row.
 *
 * Returns
 *   • Instances used by store/truck grouping, freight, and output services.
 *
 * Why / How
 *   Every later step needs a consistent normalized representation of one
 *   source item. The class stores source provenance, raw input values,
 *   casepack data, merged dimensions, and allocated weight.
 */
export class InputItem {
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ CONSTRUCTOR: constructor                                                │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Named properties from InputWorkbookReader.
   *
   * Returns
   *   • A configured InputItem instance.
   *
   * Why / How
   *   Every parsed row should use one predictable structure before any output
   *   mode applies business rules. This constructor copies shared fields and
   *   mode-specific values while giving safe defaults to optional properties.
   */
  constructor({
    sourceRow,
    customerPo,
    targetStore,
    targetStoreSourceLocation,
    itemCode = "",
    oldItemCode,
    quantity,
    pickupNumber = "",
    casepackQty = "",
    casepackValue = "",
    targetOrderCasepackQty = "",
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
    this.itemCode = itemCode;
    this.oldItemCode = oldItemCode;
    this.quantity = quantity;
    this.pickupNumber = pickupNumber;

    // LTL casepack values drive output Quantity and Total Value.
    this.casepackQty = casepackQty;
    this.casepackValue = casepackValue;
    // FTL uses the new target-order casepack quantity instead of input QTY.
    this.targetOrderCasepackQty = targetOrderCasepackQty;

    // FTL-only source values. They remain empty in LTL mode.
    this.truckNumber = truckNumber;
    this.pallet = pallet;
    this.chinaPo = chinaPo;
    this.container = container;
    this.isAdditionalPallet = isAdditionalPallet;

    // The full source G.W. value seen by this item. It is not split.
    this.grossWeight = grossWeight;
    this.grossWeightSourceLocation = grossWeightSourceLocation;

    // Merged G.W. is split across only the actual items in that G.W. range.
    this.allocatedWeight = allocatedWeight;

    // L/W/H can come from a vertically merged physical-dimension range. The
    // full dimensions are copied to every item in that physical group.
    this.length = length;
    this.width = width;
    this.height = height;
    this.dimensionSourceStartRow = dimensionSourceStartRow;
    this.dimensionSourceEndRow = dimensionSourceEndRow;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: assignDimensions                                                │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Shared dimensions and their start/end source rows.
   *
   * Returns
   *   • No returned value; updates this item's dimension fields.
   *
   * Why / How
   *   Merged L/W/H values are discovered after initial item rows are read, so
   *   the item needs a dedicated update step. The method copies the shared
   *   dimensions and records the source row range that supplied them.
   */
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
