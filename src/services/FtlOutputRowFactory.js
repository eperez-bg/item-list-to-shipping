/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: FtlOutputRowFactory.js                                               ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Builds output rows for the FTL template using Truck # grouping and         ║
 * ║ FTL-specific pickup, weight, and pallet rules.                             ║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ FTL template rules differ from LTL: grouping, pickup reference, weights,   ║
 * ║ and pallet fields are truck-based.                                         ║
 * ║                                                                            ║
 * ║ MAIN CLASSES / FUNCTIONS                                                   ║
 * ║ • FtlOutputRowFactory — FTL template rules differ from LTL: grouping,      ║
 * ║ pickup reference, weights, and pallet fields are truck-based.              ║
 * ║ • build — The FTL output must place one Order row followed immediately     ║
 * ║ by all Commodity rows for each Truck #.                                    ║
 * ║ • buildPickupReference — FTL uses the same composite value in both         ║
 * ║ Pickup Number and Primary Number, so it should be built once               ║
 * ║ consistently.                                                              ║
 * ║ • getTruckPalletCountOrError — Only the FTL Order row should receive a     ║
 * ║ truck pallet total, and missing totals need a logged visible error.        ║
 * ║ • valueOrError — Required non-dropdown FTL fields should visibly report    ║
 * ║ missing source data.                                                       ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import { OutputLoadRow } from "../models/OutputLoadRow";
import { ERROR_VALUE } from "../utils/importIssues";
import { isBlank, numericWhenPossible } from "../utils/text";

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ CLASS: FtlOutputRowFactory                                                 │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Constructor dependencies for layout, dropdowns, freight, defaults, and
 *   • issue reporting.
 *
 * Returns
 *   • A reusable FTL output row factory.
 *
 * Why / How
 *   FTL template rules differ from LTL: grouping, pickup reference, weights,
 *   and pallet fields are truck-based.
 *   It walks Truck # groups and creates rows with FTL-specific values while
 *   reusing shared dropdown and freight services.
 */
export class FtlOutputRowFactory {
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ CONSTRUCTOR: constructor                                                   │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Named dependencies from TemplateWorkbookService.
   *
   * Returns
   *   • A configured FtlOutputRowFactory.
   *
   * Why / How
   *   FTL assembly requires configuration and services but should not directly
   *   construct them itself.
   *   It stores the injected dependencies for use during build.
   */
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

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: build                                                              │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • shipment: parsed InputShipment.
   *
   * Returns
   *   • Ordered array of OutputLoadRow instances for the FTL template.
   *
   * Why / How
   *   The FTL output must place one Order row followed immediately by all
   *   Commodity rows for each Truck #.
   *   It walks source-order truck groups, calculates FTL-only fields, resolves
   *   dropdowns, and creates OutputLoadRows.
   */
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

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: buildPickupReference                                               │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • item: InputItem.
   *
   * Returns
   *   • String in the required "{Store #} Truck {Truck #} ({China PO})"
   *   • format.
   *
   * Why / How
   *   FTL uses the same composite value in both Pickup Number and Primary
   *   Number, so it should be built once consistently.
   *   It combines Store #, Truck #, and China PO; missing pieces become
   *   visible ERROR markers.
   */
  buildPickupReference(item) {
    const storeNumber = valueOrError(item.targetStore);
    const truckNumber = valueOrError(item.truckNumber);
    const chinaPo = valueOrError(item.chinaPo);

    return `${storeNumber} Truck ${truckNumber} (${chinaPo})`;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: getTruckPalletCountOrError                                         │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • truck: Truckload; item: Order-row InputItem.
   *
   * Returns
   *   • Positive pallet count or ERROR_VALUE.
   *
   * Why / How
   *   Only the FTL Order row should receive a truck pallet total, and missing
   *   totals need a logged visible error.
   *   It asks Truckload for its greatest usable pallet count and records an
   *   issue if no positive value exists.
   */
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
 *   Required non-dropdown FTL fields should visibly report missing source
 *   data.
 *   It returns a present value or the shared ERROR marker.
 */
function valueOrError(value) {
  return isBlank(value) ? ERROR_VALUE : value;
}
