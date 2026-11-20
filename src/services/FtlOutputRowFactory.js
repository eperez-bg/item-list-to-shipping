/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FtlOutputRowFactory.js                                                   ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ Converts FTL Truckloads into template rows using the NEW business rules.║
 * ║ Every Truck # is one Order/Commodity block. Repeated Item Codes on that  ║
 * ║ truck are combined FIRST, then quantity, weight, pallets, and freight     ║
 * ║ are calculated. Weight sums G.W.; value is qty × ONE casepack price.   ║
 * ║                                                                          ║
 * ║ build             — one output row per combined truck/item group.          ║
 * ║ catalogFieldsFor  — lookup description and validate Supplier Part ID.      ║
 * ║ aggregateOrError  — report invalid group metrics without stopping export. ║
 * ║ referenceFrom     — choose the first source row's CustomerPO/IF.           ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import { OutputLoadRow } from "../models/OutputLoadRow";
import { ERROR_VALUE } from "../utils/importIssues";
import { isBlank, normalizeText, numericWhenPossible } from "../utils/text";

const CONTRADICTING_CODES = "Contradicting codes";
const ADDITIONAL_PALLET = "Error: Additional Pallet";

export class FtlOutputRowFactory {
  /**
   * FtlOutputRowFactory
   * Parameters: layout, dropdownService, freightClassCalculator, defaults,
   * catalogService and issueReporter.
   * Keeps FTL business rules isolated from the older LTL row factory.
   */
  constructor({ layout, dropdownService, freightClassCalculator, defaults, catalogService, issueReporter }) {
    this.layout = layout;
    this.dropdownService = dropdownService;
    this.freightClassCalculator = freightClassCalculator;
    this.defaults = defaults;
    this.catalogService = catalogService;
    this.issueReporter = issueReporter;
  }

  /**
   * build
   * Parameters: InputShipment. Returns: ordered OutputLoadRow[] for FTL.
   * For every Truck #, combine identical Item Codes before calculating any
   * metrics. A combined group containing the truck's first raw item is Order;
   * all other combined groups are Commodity. Each group stays under its truck.
   */
  build(shipment) {
    const outputRows = [];

    shipment.trucks.forEach((truck) => {
      const combinedGroups = truck.getCombinedItemGroups();
      combinedGroups.forEach((group) => {
        const item = group.representativeItem;
        const rowNumber = this.layout.firstDataRow + outputRows.length;
        const rowType = group.isOrder ? "order" : "commodity";
        const sourceLocation = `Input row(s) ${group.sourceRows.join(", ")}, Truck # ${truck.truckNumber}`;
        const dropdown = (fieldKey, value, context) =>
          this.dropdownService.resolveExactOptionOrBlank(
            this.layout.addressFor(fieldKey, rowNumber),
            value,
            context,
          );

        // Step 1: group identical Truck # + Item Code rows (already done).
        // Step 2: sum casepack quantities AND gross weights independently.
        // Step 3: combined quantity × FIRST casepack price = Total Value.
        // Prices are never added and quantity is never used for Total Weight.
        const combinedQuantity = group.sumQuantity();
        const quantity = this.aggregateOrError(group, "Quantity", combinedQuantity);
        const totalWeight = this.aggregateOrError(group, "Total Weight", group.sumWeight());
        const totalValue = this.aggregateOrError(
          group,
          "Total Value",
          group.calculateTotalValue(combinedQuantity),
        );
        this.warnOnConflictingPrices(group);
        const length = this.aggregateOrError(group, "Length", group.largestDimension("length"));
        const width = this.aggregateOrError(group, "Width", group.largestDimension("width"));
        const height = this.aggregateOrError(group, "Height", group.largestDimension("height"));
        const freightClass = this.freightClassCalculator.calculate({
          sourceRow: group.sourceRows.join(", ") ,
          length,
          width,
          height,
          allocatedWeight: totalWeight,
        });
        const catalogFields = this.catalogFieldsFor(group);

        // The merged group's first item supplies text references; differences
        // between duplicate rows are logged rather than silently discarded.
        this.warnOnConflictingReferences(group);
        const customerPo = this.referenceFrom(item.customerPo);
        const pickupNumber = this.referenceFrom(item.pickupNumber);

        outputRows.push(new OutputLoadRow({
          type: dropdown("type", rowType, {
            field: "Type",
            sourceLocation,
          }),
          primaryNumber: customerPo,
          secondaryNumber: "",
          fixtDemandId: "",
          projectName: "",
          origin: this.dropdownService.resolveOptionContainingOrBlank(
            this.layout.addressFor("origin", rowNumber),
            this.defaults.origin,
            { field: "Origin", sourceLocation: "Application default" },
          ),
          earliestPickupDate: "",
          earliestPickupTime: "",
          latestPickupDate: "",
          latestPickupTime: "",
          pickupContactName: "",
          pickupContactEmail: "",
          pickupContactPhone: "",
          bolNumber: "",
          pickupNumber,
          originSpecialInstructions: "",
          destination: this.dropdownService.resolveDestinationOrBlank(
            this.layout.addressFor("destination", rowNumber),
            item.targetStore,
            item.targetStoreSourceLocation ?? sourceLocation,
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
          description: catalogFields.description,
          packaging: dropdown("packaging", this.defaults.packaging, {
            field: "Packaging",
            sourceLocation: "Application default",
          }),
          quantity,
          totalWeight,
          totalValue,
          freightClass: isBlank(freightClass)
            ? ""
            : numericWhenPossible(dropdown("freightClass", freightClass, {
                field: "Freight Class",
                sourceLocation,
              })),
          temperature: "",
          // ONE pallet per original input row in this combined item group.
          // Every FTL row receives its own count; Pallet Spaces is always blank.
          pallets: group.palletCount,
          palletSpaces: "",
          trailerFeet: "",
          length,
          width,
          height,
          nmfcNumber: "",
          customerPo,
          notes: "",
          sku: catalogFields.sku,
          stackable: dropdown("stackable", this.defaults.stackable, {
            field: "Stackable",
            sourceLocation: "Application default",
          }),
          expedited: "",
        }));
      });
    });

    console.info("[Fuse Order Template Filler]", {
      type: "FTL_OUTPUT_COMBINED_BY_TRUCK_AND_ITEM_CODE",
      truckCount: shipment.trucks.length,
      rawItemCount: shipment.itemCount,
      combinedOutputRowCount: outputRows.length,
    });
    return outputRows;
  }

  /**
   * catalogFieldsFor
   * Parameters: FtlItemGroup. Returns: {description, sku}.
   * Looks up the group's original Item Code. Description comes from Item
   * Description; SKU uses Supplier Part ID only when it equals the source Old
   * Item Code (case-insensitive). Otherwise SKU is 'Contradicting codes'.
   */
  catalogFieldsFor(group) {
    const representative = group.representativeItem;
    if (representative.isAdditionalPallet &&
        (isBlank(representative.itemCode) || String(representative.itemCode).startsWith("Error"))) {
      return { description: ADDITIONAL_PALLET, sku: ADDITIONAL_PALLET };
    }

    const catalogRecord = this.catalogService.lookup(representative.itemCode);
    if (!catalogRecord) {
      this.issueReporter.record({
        type: "FTL_ITEM_CODE_NOT_IN_CATALOG",
        location: `Input row ${representative.sourceRow}`,
        field: "Item Code / Description / SKU",
        message: `Item Code "${representative.itemCode}" has no entry in BEELINEGROUPFX_L3.xlsx.`,
        resolution: `Description receives ${ERROR_VALUE}; SKU receives ${CONTRADICTING_CODES}.`,
      });
      return { description: ERROR_VALUE, sku: CONTRADICTING_CODES };
    }

    let description = catalogRecord.description;
    if (isBlank(description)) {
      this.issueReporter.record({
        type: "FTL_CATALOG_DESCRIPTION_MISSING",
        location: catalogRecord.catalogLocation,
        field: "Item Description",
        message: `Item Code "${representative.itemCode}" has no Item Description.`,
        resolution: `Description receives ${ERROR_VALUE}.`,
      });
      description = ERROR_VALUE;
    }

    const inputOldCodes = group.items.map((item) => normalizeText(item.oldItemCode));
    const catalogCode = normalizeText(catalogRecord.supplierPartId);
    const codesMatch = Boolean(catalogCode) &&
      inputOldCodes.every((code) => code === catalogCode);

    if (!codesMatch) {
      this.issueReporter.record({
        type: "FTL_CONTRADICTING_SUPPLIER_PART_IDS",
        location: `Input rows ${group.sourceRows.join(", ")}; ${catalogRecord.catalogLocation}`,
        field: "SKU / Old Item Code",
        message: `Supplier Part ID "${catalogRecord.supplierPartId}" does not match every source Old Item Code (${group.items.map((item) => item.oldItemCode).join(", ")}).`,
        resolution: `SKU receives "${CONTRADICTING_CODES}". Comparison ignores capitalization.`,
      });
    }

    return {
      description,
      sku: codesMatch ? catalogRecord.supplierPartId : CONTRADICTING_CODES,
    };
  }

  /**
   * warnOnConflictingPrices
   * Parameters: FtlItemGroup. Returns: nothing; reports inconsistent prices.
   * A combined Truck/Item Code output has one casepack price. We always keep
   * the first row's original price, rather than adding prices or averaging.
   * If subsequent source rows disagree, log the chosen price and all sources.
   */
  warnOnConflictingPrices(group) {
    if (group.items.length < 2) return;

    const firstPrice = group.getCasepackPrice();
    const differingRows = group.items.slice(1).filter((item) => {
      const value = Number(String(item.casepackPrice).replace(/,/g, ""));
      return !Number.isFinite(value) || value !== firstPrice;
    });
    if (!differingRows.length) return;

    this.issueReporter.record({
      type: "FTL_COMBINED_CASEPACK_PRICE_MISMATCH",
      location: `Input rows ${group.sourceRows.join(", ")}`,
      field: "casepack price / Total Value",
      message: `Combined Item Code "${group.representativeItem.itemCode}" has different casepack prices: ${group.items.map((item) => item.casepackPrice).join(", ")}.`,
      resolution: `Using the first source row's price (${group.representativeItem.sourceRow}: ${group.representativeItem.casepackPrice}) for Total Value; prices were NOT summed.`,
    });
  }

  /**
   * aggregateOrError
   * Parameters: FtlItemGroup, field name, computed number (or null).
   * Returns: valid numeric aggregate or source-specific visible error marker.
   * Invalid arithmetic never silently becomes 0; dropdown Freight Class uses
   * a blank when metrics cannot be computed.
   */
  aggregateOrError(group, field, result) {
    if (typeof result === "number" && Number.isFinite(result)) return result;

    const isAdditionalPallet = group.items.some((item) => item.isAdditionalPallet);
    const errorText = isAdditionalPallet ? ADDITIONAL_PALLET : ERROR_VALUE;
    this.issueReporter.record({
      type: isAdditionalPallet ? "FTL_ADDITIONAL_PALLET_METRIC_MISSING" : "FTL_INVALID_AGGREGATE",
      location: `Input row(s) ${group.sourceRows.join(", ")}`,
      field,
      message: `Cannot calculate ${field} for grouped FTL Item Code "${group.representativeItem.itemCode}".`,
      resolution: `The output ${field} field receives "${errorText}". Freight Class stays blank when unavailable.`,
    });
    return errorText;
  }

  /**
   * referenceFrom
   * Parameters: one original text/value. Returns: source value or ERROR.
   * Primary Number and Pickup Number are copied directly from the first
   * matching raw row, without the old Truck/China PO combination.
   */
  referenceFrom(value) {
    return isBlank(value) ? ERROR_VALUE : value;
  }

  /**
   * warnOnConflictingReferences
   * Parameters: FtlItemGroup. Returns: nothing.
   * A duplicate Item Code may have different CustomerPO, IF, or destinations.
   * The first source row wins; logging makes this choice auditable.
   */
  warnOnConflictingReferences(group) {
    if (group.items.length < 2) return;
    const first = group.representativeItem;
    for (const [field, displayName] of [
      ["customerPo", "CustomerPO"],
      ["pickupNumber", "IF"],
      ["targetStore", "Store #"],
    ]) {
      if (group.items.some((item) => normalizeText(item[field]) !== normalizeText(first[field]))) {
        this.issueReporter.record({
          type: "FTL_COMBINED_REFERENCE_MISMATCH",
          location: `Input rows ${group.sourceRows.join(", ")}`,
          field: displayName,
          message: `Combined rows have different ${displayName} values.`,
          resolution: `Using the first source row (${first.sourceRow}) value "${first[field]}".`,
        });
      }
    }
  }
}
