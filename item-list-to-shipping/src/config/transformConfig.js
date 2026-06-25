export const APP_CONFIG = {
  template: {
    // Put the untouched original template here:
    // public/templates/fuse-order-template.xlsx
    url: "/templates/fuse-order-template.xlsx",
    outputFileName: "fuse-order-filled.xlsx",

    // Set this only if the yellow-header worksheet is not detected automatically.
    outputSheetName: null,
  },

  input: {
    // The source workbook uses fixed Excel columns, not header-name matching.
    firstPossibleDataRow: 2,
    sheetIndex: 0,
    columns: {
      // This is read separately for every input row. It supplies every
      // Target PO# column in the template.
      customerPo: "B",
      targetStore: "C",
      oldItemCode: "E",
      quantity: "F",
      length: "J",
      width: "K",
      height: "L",
      grossWeight: "M",
      pickupNumber: "T",
    },

    // A vertically merged range touching any of these columns marks one skid.
    skidGroupingColumns: ["J", "K", "L"],
  },

  defaults: {
    projectName: "Fuse Order",
    origin: "2354 Davis Ave",
    packaging: "Pallet",
    stackable: "no",
  },

  destination: {
    storeNumberLength: 4,
    optionPrefix: "Target Store #",
  },

  freight: {
    /*
      Freight class must come from your approved freight-class process.

      Set classifier to a function after you have the exact rule or table your
      company uses. It receives skid dimensions, total gross weight, and density
      in pounds per cubic foot and must return an allowed template dropdown value.

      Example only -- do not enable until confirmed by your transportation team:
      classifier: ({ densityPcf }) => (densityPcf >= 10 ? 70 : 100),
    */
    classifier: null,
  },
};

// The template has three Target PO# headers. All three receive input column B
// from the same row. "occurrence" means left-to-right occurrence in the
// template header row: B is 1, C is 2, and AN is 3.
export const TEMPLATE_FIELDS = [
  { key: "type", header: "Type" },
  { key: "customerPo", header: "Target PO#", occurrence: 1 },
  { key: "customerPoDuplicate", header: "Target PO#", occurrence: 2 },
  { key: "projectName", header: "Project Name" },
  { key: "origin", header: "Origin" },
  { key: "earliestPickupDate", header: "Earliest Pick Up Date" },
  { key: "earliestPickupTime", header: "Earliest Pick Up Time" },
  { key: "latestPickupDate", header: "Latest Pick Up Date" },
  { key: "latestPickupTime", header: "Latest Pick Up Time" },
  { key: "pickupContactName", header: "Pickup Contact Name" },
  { key: "pickupContactEmail", header: "Pickup Contact Email" },
  { key: "pickupContactPhone", header: "Pickup Contact Phone" },
  { key: "pickupNumber", header: "Pickup Number" },
  { key: "originSpecialInstructions", header: "Origin Special Instructions" },
  { key: "destination", header: "Destination" },
  { key: "earliestDeliveryDate", header: "Earliest Delivery Date" },
  { key: "earliestDeliveryTime", header: "Earliest Delivery Time" },
  { key: "latestDeliveryDate", header: "Latest Delivery Date" },
  { key: "latestDeliveryTime", header: "Latest Delivery Time" },
  { key: "deliveryContactName", header: "Delivery Contact Name" },
  { key: "deliveryContactEmail", header: "Delivery Contact Email" },
  { key: "deliveryContactPhone", header: "Delivery Contact Phone" },
  { key: "deliveryNumber", header: "Delivery Number" },
  { key: "destinationSpecialInstructions", header: "Destination Special Instructions" },
  { key: "itemId", header: "Item ID" },
  { key: "description", header: "Description" },
  { key: "packaging", header: "Packaging" },
  { key: "quantity", header: "Quantity" },
  { key: "totalWeight", header: "Total Weight" },
  { key: "totalValue", header: "Total Value" },
  { key: "freightClass", header: "Freight Class" },
  { key: "temperature", header: "Temperature" },
  { key: "pallets", header: "Pallets" },
  { key: "palletSpaces", header: "Pallet Spaces" },
  { key: "trailerFeet", header: "Trailer Feet" },
  { key: "length", header: "Length (in)" },
  { key: "width", header: "Width (in)" },
  { key: "height", header: "Height (in)" },
  { key: "nmfcNumber", header: "NMFC Number" },
  { key: "customerPoFinal", header: "Target PO#", occurrence: 3 },
  { key: "notes", header: "Notes" },
  { key: "sku", header: "SKU" },
  { key: "stackable", header: "Stackable" },
  { key: "expedited", header: "Expedited" },
];
