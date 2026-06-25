import { OutputLoadRow } from "../models/OutputLoadRow";
import { ERROR_VALUE } from "../utils/importIssues";
import {
  isBlank,
  numericWhenPossible,
  roundToDecimalPlaces,
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
      const freightClass = this.freightClassCalculator.calculate(skid);

      skid.items.forEach((item, itemIndex) => {
        const outputRowNumber = this.layout.firstDataRow + outputRows.length;
        const dropdown = (fieldKey, value, context) =>
          this.dropdownService.resolveExactOptionOrBlank(
            this.layout.addressFor(fieldKey, outputRowNumber),
            value,
            context,
          );

        const customerPo = valueOrError(item.customerPo);

        outputRows.push(
          new OutputLoadRow({
            type: dropdown("type", itemIndex === 0 ? "order" : "commodity", {
              field: "Type",
              sourceLocation: `Input row ${item.sourceRow}`,
            }),

            // Template Target PO# columns B, C, and AN all use input column B
            // from the current source row.
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
              `C${item.sourceRow}`,
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
            itemId: valueOrError(item.oldItemCode),
            description: valueOrError(item.oldItemCode),
            packaging: dropdown("packaging", this.defaults.packaging, {
              field: "Packaging",
              sourceLocation: "Application default",
            }),
            quantity: valueOrError(item.quantity),
            totalWeight: skid.perItemWeight,
            totalValue: "",
            freightClass: isBlank(freightClass)
              ? ""
              : numericWhenPossible(
                  dropdown("freightClass", freightClass, {
                    field: "Freight Class",
                    sourceLocation: `Input skid rows ${skid.sourceStartRow}-${skid.sourceEndRow}`,
                  }),
                ),
            temperature: "",

            // Template columns AG and AH accept at most three decimal places.
            pallets: roundToDecimalPlaces(skid.palletFraction, 3),
            palletSpaces: roundToDecimalPlaces(skid.palletFraction, 3),

            trailerFeet: "",
            length: valueOrError(skid.length),
            width: valueOrError(skid.width),
            height: valueOrError(skid.height),
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
