import { OutputLoadRow } from "../models/OutputLoadRow";
import { ERROR_VALUE } from "../utils/importIssues";
import {
  isBlank,
  numericWhenPossible,
} from "../utils/text";

export class OutputRowFactory {
  constructor({ layout, dropdownService, freightClassCalculator, defaults }) {
    this.layout = layout;
    this.dropdownService = dropdownService;
    this.freightClassCalculator = freightClassCalculator;
    this.defaults = defaults;
  }

  build(shipment) {
    const outputRows = [];

    shipment.skids.forEach((skid) => {
      skid.items.forEach((item) => {
        const outputRowNumber = this.layout.firstDataRow + outputRows.length;
        const purchaseOrder = shipment.getPurchaseOrderForItem(item);

        if (!purchaseOrder) {
          throw new Error(
            `Could not find a Customer PO group for input row ${item.sourceRow}.`,
          );
        }

        /*
          Type is based only on Customer PO:
          - first item for a PO = Order
          - every later item for that PO = Commodity

          It intentionally does not depend on merged dimensions or weight.
        */
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

        /*
          Pallets and Pallet Spaces are based on the physical merged L/W/H
          group, not the Customer PO. The share for all rows in this skid
          totals exactly 1.000 to three decimal places.
        */
        const palletShare = skid.getPalletShareForItem(item, 3);

        outputRows.push(
          new OutputLoadRow({
            type: dropdown("type", rowType, {
              field: "Type",
              sourceLocation: `Input row ${item.sourceRow}`,
            }),

            // Template Target PO# columns B, C, and AN all use the CustomerPO
            // value from the same input item row.
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

            // A merged G.W. is split among the actual item rows in its own
            // source G.W. range.
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

            pallets: palletShare,
            palletSpaces: palletShare,

            trailerFeet: "",

            // L/W/H values come from the item's physical merged dimension
            // group and are copied in full to every row in that group.
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

    return outputRows;
  }
}

function valueOrError(value) {
  return isBlank(value) ? ERROR_VALUE : value;
}
