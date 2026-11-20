/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: OutputRowFactory.js                                               ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                          ║
 * ║ Builds output rows for the LTL template using Store # based truck       ║
 * ║ groups, then combines duplicate items inside each truck before writing  ║
 * ║ the final template rows.                                                ║
 * ║                                                                         ║
 * ║ WHY THIS FILE EXISTS                                                    ║
 * ║ LTL business rules now have multiple layers: rows are first grouped by  ║
 * ║ Store # into generated trucks, then repeated Item Code + CustomerPO     ║
 * ║ rows inside each truck are combined. Keeping this logic separate from   ║
 * ║ XML/template code makes the transformation easier to debug and change.  ║
 * ║                                                                         ║
 * ║ MAIN CLASSES / FUNCTIONS                                                ║
 * ║ • OutputRowFactory — Converts LTL StoreTruck groups into output rows.   ║
 * ║ • build — Writes grouped/combined Order and Commodity rows.             ║
 * ║ • buildCombinedRowValues — Aggregates duplicate item rows for output.   ║
 * ║ • calculateCombinedCasepackQuantity — Sums carton quantities.           ║
 * ║ • calculateCombinedTotalValue — Sums item-level total values.           ║
 * ║ • calculateCombinedTotalWeight — Sums allocated item weights.           ║
 * ║ • calculateCombinedPalletShare — Sums merged-dimension pallet shares.   ║
 * ║ • getCombinedDimensions — Keeps the largest L/W/H values in a group.    ║
 * ║ • createSkidIndex — Finds physical merged groups for pallet shares.     ║
 * ║ • valueOrError — Converts missing normal-field values to ERROR.         ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import { OutputLoadRow } from "../models/OutputLoadRow";
import { ERROR_VALUE } from "../utils/importIssues";
import {
  isBlank,
  numericWhenPossible,
  roundToDecimalPlaces,
  toFiniteNumber,
} from "../utils/text";

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ CLASS: OutputRowFactory                                                 │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Constructor dependencies for layout, dropdowns, freight, defaults, and
 *     optional issue reporting.
 *
 * Returns
 *   • A reusable LTL output row factory.
 *
 * Why / How
 *   LTL output needs store/truck-level values and item-level values in the same
 *   row. The factory walks StoreTruck groups, combines matching source rows,
 *   computes aggregate quantities/weights/pallets/dimensions, and emits one
 *   OutputLoadRow for each combined item group.
 */
export class OutputRowFactory {
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ CONSTRUCTOR: constructor                                                │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Named service dependencies from TemplateWorkbookService.
   *
   * Returns
   *   • A configured OutputRowFactory.
   */
  constructor({
    layout,
    dropdownService,
    freightClassCalculator,
    defaults,
    issueReporter = null,
  }) {
    this.layout = layout;
    this.dropdownService = dropdownService;
    this.freightClassCalculator = freightClassCalculator;
    this.defaults = defaults;
    this.issueReporter = issueReporter;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: build                                                           │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • shipment: parsed InputShipment.
   *
   * Returns
   *   • Ordered array of OutputLoadRow instances for the LTL template.
   *
   * Why / How
   *   The downloaded LTL worksheet should scroll as grouped trucks: an Order
   *   row for a Store # followed by related Commodity rows. Inside each truck,
   *   duplicate Item Code + CustomerPO rows are combined before output so the
   *   resulting sheet has fewer, cleaner item rows.
   */
  build(shipment) {
    const outputRows = [];
    const skidByItem = this.createSkidIndex(shipment);

    shipment.storeTrucks.forEach((truck) => {
      truck.getCombinedItemGroups().forEach((combinedGroup) => {
        const outputRowNumber = this.layout.firstDataRow + outputRows.length;
        const rowValues = this.buildCombinedRowValues({
          truck,
          combinedGroup,
          outputRowNumber,
          skidByItem,
        });

        outputRows.push(new OutputLoadRow(rowValues));
      });
    });

    console.info("[Fuse Order Template Filler]", {
      type: "LTL_OUTPUT_GROUPED_BY_STORE_TRUCK_AND_COMBINED_ITEMS",
      storeCount: shipment.storeCount,
      truckCount: shipment.storeTruckCount,
      rawItemCount: shipment.itemCount,
      outputRowCount: outputRows.length,
    });

    return outputRows;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: buildCombinedRowValues                                          │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • truck: StoreTruck containing the raw rows for one Store #.
   *   • combinedGroup: StoreTruckItemGroup representing one output row.
   *   • outputRowNumber: worksheet row being filled.
   *   • skidByItem: Map<InputItem, InputSkid> for pallet-share lookup.
   *
   * Returns
   *   • Plain object of template field values used by OutputLoadRow.
   *
   * Why / How
   *   This is the main LTL mapping step. It keeps group-level values from the
   *   raw Order item, uses the first item in the combined group for normal
   *   item-level fields, and uses aggregate helpers for quantity, weight,
   *   pallet share, total value, and dimensions.
   */
  buildCombinedRowValues({ truck, combinedGroup, outputRowNumber, skidByItem }) {
    const item = combinedGroup.representativeItem;
    const isOrderRow = truck.isOrderGroup(combinedGroup);
    const rowType = isOrderRow ? "order" : "commodity";
    const groupTargetPo = valueOrError(truck.targetPo);
    const groupPickupNumber = valueOrError(truck.pickupNumber);
    const quantityDetails = this.calculateCombinedCasepackQuantity(combinedGroup);
    const calculatedQuantity = quantityDetails.total;
    const totalValue = this.calculateCombinedTotalValue(
      combinedGroup,
      quantityDetails,
    );
    const totalWeight = this.calculateCombinedTotalWeight(combinedGroup);
    const palletShare = this.calculateCombinedPalletShare(
      combinedGroup,
      skidByItem,
    );
    const dimensions = this.getCombinedDimensions(combinedGroup);
    const freightClass = this.freightClassCalculator.calculate({
      ...item,
      sourceRow: combinedGroup.sourceRows,
      allocatedWeight: totalWeight,
      length: dimensions.length,
      width: dimensions.width,
      height: dimensions.height,
      dimensionSourceStartRow: null,
      dimensionSourceEndRow: null,
    });

    const dropdown = (fieldKey, value, context) =>
      this.dropdownService.resolveExactOptionOrBlank(
        this.layout.addressFor(fieldKey, outputRowNumber),
        value,
        context,
      );

    return {
      type: dropdown("type", rowType, {
        field: "Type",
        sourceLocation:
          `Input rows ${combinedGroup.sourceRows}, Store # ` +
          `${valueOrError(item.targetStore)}`,
      }),

      // The first two Target PO# columns are group-level fields. They use the
      // raw Order item's CustomerPO for every row in the same Store group.
      customerPo: groupTargetPo,
      customerPoDuplicate: groupTargetPo,

      projectName: this.defaults.projectName,
      origin: this.dropdownService.resolveOptionContainingOrBlank(
        this.layout.addressFor("origin", outputRowNumber),
        this.defaults.origin,
        {
          field: "Origin",
          sourceLocation: "Application default",
        },
      ),

      earliestPickupDate: "",
      earliestPickupTime: "",
      latestPickupDate: "",
      latestPickupTime: "",
      pickupContactName: "",
      pickupContactEmail: "",
      pickupContactPhone: "",

      // Pickup Number is group-level from the raw Order row's IF.
      pickupNumber: groupPickupNumber,

      // Origin Special Instructions keeps the first IF number from this
      // combined item group. Matching rows with different IF values still
      // combine, but the first IF is intentionally retained.
      originSpecialInstructions: valueOrError(item.pickupNumber),

      destination: this.dropdownService.resolveDestinationOrBlank(
        this.layout.addressFor("destination", outputRowNumber),
        item.targetStore,
        item.targetStoreSourceLocation ?? `Input row ${item.sourceRow}`,
      ),

      earliestDeliveryDate: "",
      earliestDeliveryTime: "",
      latestDeliveryDate: "",
      latestDeliveryTime: "",
      deliveryContactName: "",
      deliveryContactEmail: "",
      deliveryContactPhone: "",
      deliveryNumber: "",
      destinationSpecialInstructions: "",

      itemId: valueOrError(item.itemCode),
      description: valueOrError(item.oldItemCode),
      packaging: dropdown("packaging", this.defaults.ltlPackaging, {
        field: "Packaging",
        sourceLocation: "Application default",
      }),
      quantity: calculatedQuantity,
      totalWeight,
      totalValue,

      freightClass: isBlank(freightClass)
        ? ""
        : numericWhenPossible(
            dropdown("freightClass", freightClass, {
              field: "Freight Class",
              sourceLocation: `Input rows ${combinedGroup.sourceRows}`,
            }),
          ),

      temperature: "",
      pallets: palletShare,
      palletSpaces: palletShare,
      trailerFeet: "",
      length: valueOrError(dimensions.length),
      width: valueOrError(dimensions.width),
      height: valueOrError(dimensions.height),
      nmfcNumber: "",

      // The PO Number column is item-level. Because combined rows require the
      // same CustomerPO to combine, the representative item's CustomerPO is the
      // CustomerPO for the whole combined item row.
      customerPoFinal: valueOrError(item.customerPo),
      notes: "",
      sku: valueOrError(item.oldItemCode),
      stackable: dropdown("stackable", this.defaults.stackable, {
        field: "Stackable",
        sourceLocation: "Application default",
      }),
      expedited: "",
    };
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: calculateCombinedCasepackQuantity                               │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • combinedGroup: StoreTruckItemGroup.
   *
   * Returns
   *   • Object with total and per-item quantity results.
   *
   * Why / How
   *   Each raw item first follows the normal LTL quantity rule: QTY divided by
   *   casepack qty. Matching item rows are then combined by summing those
   *   already-calculated carton quantities.
   */
  calculateCombinedCasepackQuantity(combinedGroup) {
    const itemResults = combinedGroup.items.map((item) => ({
      item,
      quantity: this.calculateCasepackQuantity(item),
    }));

    const total = this.sumCalculatedValues({
      values: itemResults.map((result) => result.quantity),
      decimalPlaces: 3,
      issue: {
        type: "LTL_COMBINED_QUANTITY_UNAVAILABLE",
        location: `Input rows ${combinedGroup.sourceRows}`,
        field: "Quantity",
        message:
          "Could not combine Quantity because one or more matching source rows " +
          "could not calculate QTY / casepack qty.",
        resolution: 'The combined Quantity output cell will receive "ERROR".',
      },
    });

    return {
      total,
      itemResults,
    };
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: calculateCasepackQuantity                                       │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • item: InputItem.
   *
   * Returns
   *   • QTY / casepack qty rounded to three decimals, or ERROR text.
   *
   * Why / How
   *   The LTL template wants carton quantity instead of raw unit quantity. This
   *   method validates both source numbers and divides QTY by casepack qty.
   */
  calculateCasepackQuantity(item) {
    const rawQuantity = toFiniteNumber(item.quantity);
    const casepackQty = toFiniteNumber(item.casepackQty);

    if (rawQuantity !== null && casepackQty !== null && casepackQty > 0) {
      return roundToDecimalPlaces(rawQuantity / casepackQty, 3);
    }

    this.recordIssue({
      type: "LTL_CASEPACK_QUANTITY_UNAVAILABLE",
      location: `Input row ${item.sourceRow}`,
      field: "Quantity",
      message:
        `Could not calculate Quantity from QTY "${item.quantity}" and ` +
        `casepack qty "${item.casepackQty}".`,
      resolution: 'The Quantity output cell will receive "ERROR".',
    });

    return ERROR_VALUE;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: calculateCombinedTotalValue                                     │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • combinedGroup: StoreTruckItemGroup.
   *   • quantityDetails: result from calculateCombinedCasepackQuantity.
   *
   * Returns
   *   • Sum of item-level total values rounded to two decimals, or ERROR.
   *
   * Why / How
   *   Combining should preserve value from every source row. The method first
   *   calculates each source row's total value from its own carton quantity and
   *   casepack Value, then sums those item-level totals.
   */
  calculateCombinedTotalValue(combinedGroup, quantityDetails) {
    const values = quantityDetails.itemResults.map(({ item, quantity }) =>
      this.calculateTotalValue(item, quantity),
    );

    return this.sumCalculatedValues({
      values,
      decimalPlaces: 2,
      issue: {
        type: "LTL_COMBINED_TOTAL_VALUE_UNAVAILABLE",
        location: `Input rows ${combinedGroup.sourceRows}`,
        field: "Total Value",
        message:
          "Could not combine Total Value because one or more matching source " +
          "rows could not calculate Quantity × casepack Value.",
        resolution: 'The combined Total Value output cell will receive "ERROR".',
      },
    });
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: calculateTotalValue                                             │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • item: InputItem; calculatedQuantity from calculateCasepackQuantity.
   *
   * Returns
   *   • calculatedQuantity × casepack Value rounded to two decimals, or ERROR.
   *
   * Why / How
   *   Total Value depends on the already converted carton quantity. Reusing the
   *   calculated quantity keeps the output consistent with the Quantity column.
   */
  calculateTotalValue(item, calculatedQuantity) {
    const quantity = toFiniteNumber(calculatedQuantity);
    const casepackValue = toFiniteNumber(item.casepackValue);

    if (quantity !== null && casepackValue !== null) {
      return roundToDecimalPlaces(quantity * casepackValue, 2);
    }

    this.recordIssue({
      type: "LTL_TOTAL_VALUE_UNAVAILABLE",
      location: `Input row ${item.sourceRow}`,
      field: "Total Value",
      message:
        `Could not calculate Total Value from Quantity "${calculatedQuantity}" ` +
        `and casepack Value "${item.casepackValue}".`,
      resolution: 'The Total Value output cell will receive "ERROR".',
    });

    return ERROR_VALUE;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: calculateCombinedTotalWeight                                    │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • combinedGroup: StoreTruckItemGroup.
   *
   * Returns
   *   • Sum of allocated weights for the combined rows, or ERROR.
   *
   * Why / How
   *   Weight was already split from merged G.W. ranges during input parsing.
   *   Combining duplicates should add those allocated item weights together.
   */
  calculateCombinedTotalWeight(combinedGroup) {
    return this.sumCalculatedValues({
      values: combinedGroup.items.map((item) => item.allocatedWeight),
      decimalPlaces: 4,
      issue: {
        type: "LTL_COMBINED_WEIGHT_UNAVAILABLE",
        location: `Input rows ${combinedGroup.sourceRows}`,
        field: "Total Weight",
        message:
          "Could not combine Total Weight because one or more matching source " +
          "rows had an invalid allocated G.W. value.",
        resolution: 'The combined Total Weight output cell will receive "ERROR".',
      },
    });
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: calculateCombinedPalletShare                                    │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • combinedGroup: StoreTruckItemGroup.
   *   • skidByItem: Map<InputItem, InputSkid>.
   *
   * Returns
   *   • Sum of pallet/pallet-space shares rounded to three decimals, or ERROR.
   *
   * Why / How
   *   Pallets and Pallet Spaces still come from physical merged L/W/H groups.
   *   When duplicate items combine, their already-assigned pallet fractions are
   *   added together so the combined row represents all source rows.
   */
  calculateCombinedPalletShare(combinedGroup, skidByItem) {
    const shares = combinedGroup.items.map((item) => {
      const skid = skidByItem.get(item);

      if (!skid) {
        this.recordIssue({
          type: "LTL_PALLET_SHARE_SKID_MISSING",
          location: `Input row ${item.sourceRow}`,
          field: "Pallets / Pallet Spaces",
          message:
            "Could not find the merged-dimension group for this item while " +
            "combining duplicate rows.",
          resolution:
            'The combined Pallets and Pallet Spaces output cells will receive "ERROR".',
        });

        return ERROR_VALUE;
      }

      return skid.getPalletShareForItem(item, 3);
    });

    return this.sumCalculatedValues({
      values: shares,
      decimalPlaces: 3,
      issue: {
        type: "LTL_COMBINED_PALLET_SHARE_UNAVAILABLE",
        location: `Input rows ${combinedGroup.sourceRows}`,
        field: "Pallets / Pallet Spaces",
        message:
          "Could not combine pallet shares because one or more matching source " +
          "rows had an invalid pallet share.",
        resolution:
          'The combined Pallets and Pallet Spaces output cells will receive "ERROR".',
      },
    });
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getCombinedDimensions                                           │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • combinedGroup: StoreTruckItemGroup.
   *
   * Returns
   *   • Object containing the largest valid length, width, and height.
   *
   * Why / How
   *   When duplicate items have different dimensions, management wants the
   *   combined output row to keep the largest L/W/H values. This is the safer
   *   assumption for freight and template reporting.
   */
  getCombinedDimensions(combinedGroup) {
    return {
      length: this.getMaximumNumericField(combinedGroup, "length", "Length"),
      width: this.getMaximumNumericField(combinedGroup, "width", "Width"),
      height: this.getMaximumNumericField(combinedGroup, "height", "Height"),
    };
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getMaximumNumericField                                          │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • combinedGroup: StoreTruckItemGroup.
   *   • fieldKey: InputItem numeric field name.
   *   • fieldLabel: readable label for issue logging.
   *
   * Returns
   *   • Largest finite value, or ERROR when no valid value exists.
   *
   * Why / How
   *   Some combined items can pull dimensions from different physical groups.
   *   The method ignores invalid values when at least one valid value exists,
   *   but logs an issue if the whole combined group lacks a usable number.
   */
  getMaximumNumericField(combinedGroup, fieldKey, fieldLabel) {
    const values = combinedGroup.items
      .map((item) => toFiniteNumber(item[fieldKey]))
      .filter((value) => value !== null);

    if (values.length > 0) {
      return Math.max(...values);
    }

    this.recordIssue({
      type: "LTL_COMBINED_DIMENSION_UNAVAILABLE",
      location: `Input rows ${combinedGroup.sourceRows}`,
      field: fieldLabel,
      message: `Could not choose a maximum ${fieldLabel} because no matching source row had a valid value.`,
      resolution: `The ${fieldLabel} output cell will receive "ERROR".`,
    });

    return ERROR_VALUE;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: sumCalculatedValues                                             │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • values: values to sum.
   *   • decimalPlaces: final rounding precision.
   *   • issue: issue to record if any value is invalid.
   *
   * Returns
   *   • Rounded numeric sum, or ERROR.
   *
   * Why / How
   *   Several combine rules are just "sum all valid row-level calculations".
   *   This helper keeps that behavior consistent and prevents partial sums from
   *   hiding a source row that failed calculation.
   */
  sumCalculatedValues({ values, decimalPlaces, issue }) {
    const numericValues = values.map((value) => toFiniteNumber(value));

    if (numericValues.some((value) => value === null)) {
      this.recordIssue(issue);
      return ERROR_VALUE;
    }

    return roundToDecimalPlaces(
      numericValues.reduce((sum, value) => sum + value, 0),
      decimalPlaces,
    );
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: createSkidIndex                                                 │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • shipment: parsed InputShipment.
   *
   * Returns
   *   • Map<InputItem, InputSkid>.
   *
   * Why / How
   *   LTL output is now ordered by Store # and duplicate item groups, but
   *   pallet shares are still tied to the physical merged L/W/H group. This
   *   index links each raw item back to that physical group.
   */
  createSkidIndex(shipment) {
    const skidByItem = new Map();

    shipment.skids.forEach((skid) => {
      skid.items.forEach((item) => skidByItem.set(item, skid));
    });

    return skidByItem;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: recordIssue                                                     │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • issue: structured issue object.
   *
   * Returns
   *   • No returned value.
   *
   * Why / How
   *   Calculation problems should show in the browser console and be counted in
   *   the UI status. The reporter handles both when it is available.
   */
  recordIssue(issue) {
    if (this.issueReporter) {
      this.issueReporter.record(issue);
      return;
    }

    console.error("[Fuse Order Template Filler]", issue);
  }
}

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: valueOrError                                                   │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • value: any source/output value.
 *
 * Returns
 *   • Original value or ERROR_VALUE.
 *
 * Why / How
 *   Required non-dropdown fields should visibly report missing source data.
 */
function valueOrError(value) {
  return isBlank(value) ? ERROR_VALUE : value;
}
