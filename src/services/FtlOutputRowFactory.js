import { OutputLoadRow } from "../models/OutputLoadRow";
import { ERROR_VALUE } from "../utils/importIssues";
import { isBlank, numericWhenPossible } from "../utils/text";

/*
  FTL output factory.

  The completed template is grouped by Truck # in first-appearance order:
  every Truck's Order row is immediately followed by its Commodity rows.
*/
export class FtlOutputRowFactory {
  constructor({
    layout,
    dropdownService,
    freightClassCalculator,
    defaults,
    issueReporter,
  }) {
    this.layout = layout;
    this.dropdownService = dropdownService;
    this.freightClassCalculator = freightClassCalculator;
    this.defaults = defaults;
    this.issueReporter = issueReporter;
  }

  build(shipment) {
    const outputRows = [];

    /*
      InputShipment creates trucks in the order each Truck # first appears in
      the source file. Walking the truck groups directly keeps all related
      FTL rows together in the downloaded template.
    */
    shipment.trucks.forEach((truck) => {
      truck.items.forEach((item) => {
        const outputRowNumber = this.layout.firstDataRow + outputRows.length;
        const isOrderRow = truck.isFirstItem(item);
        const rowType = isOrderRow ? "order" : "commodity";
        const pickupReference = this.buildPickupReference(item);
        const freightClass = this.freightClassCalculator.calculate(item);
        const truckPalletCount = isOrderRow
          ? this.getTruckPalletCountOrError(truck, item)
          : "";

        const dropdown = (fieldKey, value, context) =>
          this.dropdownService.resolveExactOptionOrBlank(
            this.layout.addressFor(fieldKey, outputRowNumber),
            value,
            context,
          );

        outputRows.push(
          new OutputLoadRow({
            type: dropdown("type", rowType, {
              field: "Type",
              sourceLocation:
                `Input row ${item.sourceRow}, Truck # ` +
                `${valueOrError(item.truckNumber)}`,
            }),
            primaryNumber: pickupReference,
            secondaryNumber: "",
            fixtDemandId: "",
            projectName: "",

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
            bolNumber: "",
            pickupNumber: pickupReference,
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

            // FTL: the Truck's Order row keeps its full source G.W.; every
            // Commodity row deliberately receives 1.
            totalWeight: isOrderRow ? valueOrError(item.grossWeight) : 1,
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

            // Only the Order row receives the truck's greatest Pallet count.
            pallets: truckPalletCount,
            palletSpaces: truckPalletCount,

            trailerFeet: "",
            length: valueOrError(item.length),
            width: valueOrError(item.width),
            height: valueOrError(item.height),
            nmfcNumber: "",
            customerPo: valueOrError(item.customerPo),
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
      type: "FTL_OUTPUT_GROUPED_BY_TRUCK",
      truckCount: shipment.trucks.length,
      outputRowCount: outputRows.length,
    });

    return outputRows;
  }

  buildPickupReference(item) {
    const storeNumber = valueOrError(item.targetStore);
    const truckNumber = valueOrError(item.truckNumber);
    const chinaPo = valueOrError(item.chinaPo);

    return `${storeNumber} Truck ${truckNumber} (${chinaPo})`;
  }

  getTruckPalletCountOrError(truck, item) {
    const palletCount = truck.getGreatestPalletCount();

    if (palletCount !== ERROR_VALUE) {
      return palletCount;
    }

    this.issueReporter.record({
      type: "FTL_TRUCK_PALLET_COUNT_UNAVAILABLE",
      location: `Input Truck # ${valueOrError(item.truckNumber)}`,
      field: "Pallets / Pallet Spaces",
      message:
        `Truck # ${valueOrError(item.truckNumber)} has no usable positive ` +
        "Pallet value.",
      resolution:
        'The FTL Order row will receive "ERROR" in Pallets and Pallet Spaces.',
    });

    return ERROR_VALUE;
  }
}

function valueOrError(value) {
  return isBlank(value) ? ERROR_VALUE : value;
}
