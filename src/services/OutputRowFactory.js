/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: OutputRowFactory.js                                                  ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Builds output rows for the LTL template using Customer PO grouping while   ║
 * ║ retaining physical merged-cell rules for pallet shares.                    ║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ LTL has its own PO-based grouping rules, while template field assembly     ║
 * ║ should stay separate from UI and XML code.                                 ║
 * ║                                                                            ║
 * ║ MAIN CLASSES / FUNCTIONS                                                   ║
 * ║ • OutputRowFactory — LTL has its own PO-based grouping rules, while        ║
 * ║ template field assembly should stay separate from UI and XML code.         ║
 * ║ • build — The downloaded LTL worksheet must group all same-PO rows         ║
 * ║ together even when their items came from different merged-cell groups.     ║
 * ║ • createSkidIndex — PO grouping controls row order, but pallet shares      ║
 * ║ still depend on the physical L/W/H group.                                  ║
 * ║ • valueOrError — Normal output cells should make a missing source value    ║
 * ║ visible rather than silently sending an empty value.                       ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import { OutputLoadRow } from "../models/OutputLoadRow";
import { ERROR_VALUE } from "../utils/importIssues";
import { isBlank, numericWhenPossible } from "../utils/text";

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ CLASS: OutputRowFactory                                                    │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Constructor dependencies for layout, dropdowns, freight, and defaults.
 *
 * Returns
 *   • A reusable LTL output row factory.
 *
 * Why / How
 *   LTL has its own PO-based grouping rules, while template field assembly
 *   should stay separate from UI and XML code.
 *   It walks PO groups in first-appearance order and maps every item into an
 *   OutputLoadRow using shared dropdown and freight services.
 */
export class OutputRowFactory {
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ CONSTRUCTOR: constructor                                                   │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Required service dependencies.
   *
   * Returns
   *   • A configured OutputRowFactory.
   *
   * Why / How
   *   LTL row creation depends on several collaborators, which are injected to
   *   keep the factory testable and focused.
   *   It saves the layout, validation service, freight calculator, and static
   *   defaults.
   */
  constructor({ layout, dropdownService, freightClassCalculator, defaults }) {
    this.layout = layout;
    this.dropdownService = dropdownService;
    this.freightClassCalculator = freightClassCalculator;
    this.defaults = defaults;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: build                                                              │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • shipment: parsed InputShipment.
   *
   * Returns
   *   • Ordered array of OutputLoadRow instances for the LTL template.
   *
   * Why / How
   *   The downloaded LTL worksheet must group all same-PO rows together even
   *   when their items came from different merged-cell groups.
   *   It builds a skid lookup, walks PO groups in source-order, determines
   *   Order/Commodity, calculates freight/pallet fields, and creates
   *   OutputLoadRows.
   */
  build(shipment) {
    const outputRows = [];
    const skidByItem = this.createSkidIndex(shipment);

    /*
      InputShipment creates purchaseOrders in the order each PO first appears
      in the source file. Walking each PO group directly guarantees that the
      Order row and all of its Commodity rows stay together in the download.
    */
    shipment.purchaseOrders.forEach((purchaseOrder) => {
      purchaseOrder.items.forEach((item) => {
        const skid = skidByItem.get(item);

        if (!skid) {
          throw new Error(
            `Could not find the merged-dimension group for input row ${item.sourceRow}.`,
          );
        }

        const outputRowNumber = this.layout.firstDataRow + outputRows.length;
        const rowType = purchaseOrder.isFirstItem(item)
          ? "order"
          : "commodity";

        const dropdown = (fieldKey, value, context) =>
          this.dropdownService.resolveExactOptionOrBlank(
            this.layout.addressFor(fieldKey, outputRowNumber),
            value,
            context,
          );

        const customerPo = valueOrError(item.customerPo);
        const freightClass = this.freightClassCalculator.calculate(item);
        const palletShare = skid.getPalletShareForItem(item, 3);

        outputRows.push(
          new OutputLoadRow({
            type: dropdown("type", rowType, {
              field: "Type",
              sourceLocation: `Input row ${item.sourceRow}`,
            }),

            // All three LTL Target PO# columns use this row's CustomerPO.
            customerPo,
            customerPoDuplicate: customerPo,
            customerPoFinal: customerPo,

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
            pickupNumber: valueOrError(item.pickupNumber),
            originSpecialInstructions: "",

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

            itemId: "",
            description: valueOrError(item.oldItemCode),
            packaging: dropdown("packaging", this.defaults.packaging, {
              field: "Packaging",
              sourceLocation: "Application default",
            }),
            quantity: valueOrError(item.quantity),
            totalWeight: valueOrError(item.allocatedWeight),
            totalValue: "",

            freightClass: isBlank(freightClass)
              ? ""
              : numericWhenPossible(
                  dropdown("freightClass", freightClass, {
                    field: "Freight Class",
                    sourceLocation: `Input row ${item.sourceRow}`,
                  }),
                ),

            temperature: "",

            // This remains based on the item's physical merged L/W/H group,
            // not on its Customer PO.
            pallets: palletShare,
            palletSpaces: palletShare,

            trailerFeet: "",
            length: valueOrError(item.length),
            width: valueOrError(item.width),
            height: valueOrError(item.height),
            nmfcNumber: "",
            notes: "",
            sku: valueOrError(item.oldItemCode),
            stackable: dropdown("stackable", this.defaults.stackable, {
              field: "Stackable",
              sourceLocation: "Application default",
            }),
            expedited: "",
          }),
        );
      });
    });

    console.info("[Fuse Order Template Filler]", {
      type: "LTL_OUTPUT_GROUPED_BY_CUSTOMER_PO",
      purchaseOrderCount: shipment.purchaseOrders.length,
      outputRowCount: outputRows.length,
    });

    return outputRows;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: createSkidIndex                                                    │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • shipment: InputShipment.
   *
   * Returns
   *   • Map<InputItem, InputSkid>.
   *
   * Why / How
   *   PO grouping controls row order, but pallet shares still depend on the
   *   physical L/W/H group.
   *   It builds a Map from each item to its InputSkid.
   */
  createSkidIndex(shipment) {
    const skidByItem = new Map();

    shipment.skids.forEach((skid) => {
      skid.items.forEach((item) => {
        skidByItem.set(item, skid);
      });
    });

    return skidByItem;
  }
}

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: valueOrError                                                     │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • value: any source/output value.
 *
 * Returns
 *   • Original value or ERROR_VALUE.
 *
 * Why / How
 *   Normal output cells should make a missing source value visible rather
 *   than silently sending an empty value.
 *   It returns the value when present or the shared ERROR marker when blank.
 */
function valueOrError(value) {
  return isBlank(value) ? ERROR_VALUE : value;
}
