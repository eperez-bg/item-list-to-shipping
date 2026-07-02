import { OutputLoadRow } from "../models/OutputLoadRow";
import { ERROR_VALUE } from "../utils/importIssues";
import { isBlank, numericWhenPossible } from "../utils/text";

/*
  LTL output factory.

  The completed template is grouped by Customer PO in first-appearance order:
  every PO's Order row is immediately followed by its Commodity rows.

  Physical merged L/W/H groups are still used only for Pallets / Pallet Spaces.
*/
export class OutputRowFactory {
  constructor({ layout, dropdownService, freightClassCalculator, defaults }) {
    this.layout = layout;
    this.dropdownService = dropdownService;
    this.freightClassCalculator = freightClassCalculator;
    this.defaults = defaults;
  }

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

function valueOrError(value) {
  return isBlank(value) ? ERROR_VALUE : value;
}
