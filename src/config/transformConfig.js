/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: transformConfig.js                                                   ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Centralizes the fixed business rules, header names, template locations,    ║
 * ║ defaults, catalog locations, and field-to-template mappings for both modes.║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ Keeping this responsibility in one focused file makes the spreadsheet      ║
 * ║ workflow easier to understand, change, and test without affecting          ║
 * ║ unrelated code.                                                            ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

export const SHIPPING_MODES = Object.freeze({
  LTL: "LTL",
  FTL: "FTL",
});

const templateBaseUrl = import.meta.env.BASE_URL;

export const APP_CONFIG = {
  templates: {
    [SHIPPING_MODES.LTL]: {
      // Untouched original template:
      // public/templates/fuse-order-template.xlsx
      url: `${templateBaseUrl}templates/fuse-order-template.xlsx`,
      outputFileName: "fuse-order-filled.xlsx",
      outputSheetName: null,
    },

    [SHIPPING_MODES.FTL]: {
      // Untouched original truckload template:
      // public/templates/fuse-order-ftl-template.xlsx
      url: `${templateBaseUrl}templates/fuse-order-ftl-template.xlsx`,
      outputFileName: "fuse-order-ftl-filled.xlsx",
      outputSheetName: "C3257529-Truckload",
    },
  },

  input: {
    sheetIndex: 0,

    /*
      Input headers are exact except for capitalization. The Excel column letters
      may change when unused columns are deleted; the app locates the header row
      and uses the matching columns at runtime.
    */
    headers: {
      customerPo: "CustomerPO",
      targetStore: "Store #",
      itemCode: "Item Code",
      oldItemCode: "Old Item code",
      quantity: "QTY",
      length: 'L"',
      width: 'W"',
      height: 'H"',
      grossWeight: "G.W. (LBS)",
      pickupNumber: "IF",
      casepackQty: "casepack qty",
      casepackValue: "casepack Value",
      targetOrderCasepackQty: "target order casepack qty",
      truckNumber: "Truck #",
      pallet: "Pallet",
      chinaPo: "China PO",
      container: "container",
    },

    modes: {
      [SHIPPING_MODES.LTL]: {
        requiredFields: [
          "customerPo",
          "targetStore",
          "itemCode",
          "oldItemCode",
          "quantity",
          "length",
          "width",
          "height",
          "grossWeight",
          "pickupNumber",
          "casepackQty",
          "casepackValue",
        ],
        skidGroupingFields: ["length", "width", "height"],
        optionalFields: ["container"],
      },

      [SHIPPING_MODES.FTL]: {
        // FTL now uses Item Code + target order casepack qty, not source QTY.
        // Pallet and China PO are accepted if present but are no longer used.
        requiredFields: [
          "customerPo",
          "targetStore",
          "itemCode",
          "oldItemCode",
          "targetOrderCasepackQty",
          "length",
          "width",
          "height",
          "grossWeight",
          "pickupNumber",
          "truckNumber",
        ],
        skidGroupingFields: ["length", "width", "height"],
        optionalFields: ["container", "pallet", "chinaPo"],
      },
    },
  },

  defaults: {
    projectName: "Fuse Order",
    origin: "2354 Davis Ave",
    packaging: "Pallet",
    ltlPackaging: "Cartons",
    stackable: "no",
  },

  destination: {
    storeNumberLength: 4,
    optionPrefix: "Target Store #",
  },

  freight: {
    // Leave null to use the automatic density bands already used by the LTL flow.
    classifier: null,
  },

  // The catalog is a static browser asset; only FTL loads it during export.
  // Place BEELINEGROUPFX_L3.xlsx at public/catalogs/BEELINEGROUPFX_L3.xlsx.
  ftlCatalog: {
    url: `${templateBaseUrl}catalogs/BEELINEGROUPFX_L3.xlsx`,
    // Catalog labels are on Excel row 11 (one-based), not row 1.
    headerRowNumber: 11,
    headers: {
      // Input Item Code matches this catalog field; the catalog does not
      // have a separate "Item Code" column.
      manufacturerPartId: "Manufacturer Part ID",
      description: "Item Description",
      supplierPartId: "Supplier Part ID",
    },
  },
};

export const TEMPLATE_FIELDS_BY_MODE = {
  [SHIPPING_MODES.LTL]: [
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
    { key: "customerPoFinal", header: "PO Number: 10 digit 450, 700" },
    { key: "notes", header: "Notes" },
    { key: "sku", header: "SKU" },
    { key: "stackable", header: "Stackable" },
    { key: "expedited", header: "Expedited" },
  ],

  [SHIPPING_MODES.FTL]: [
    { key: "type", header: "Type" },
    { key: "primaryNumber", header: "Primary Number" },
    { key: "secondaryNumber", header: "Secondary Number" },
    { key: "fixtDemandId", header: "FIXT Demand ID (required for VDN)" },
    { key: "projectName", header: "Project Name" },
    { key: "origin", header: "Origin" },
    { key: "earliestPickupDate", header: "Earliest Pick Up Date" },
    { key: "earliestPickupTime", header: "Earliest Pick Up Time" },
    { key: "latestPickupDate", header: "Latest Pick Up Date" },
    { key: "latestPickupTime", header: "Latest Pick Up Time" },
    { key: "pickupContactName", header: "Pickup Contact Name" },
    { key: "pickupContactEmail", header: "Pickup Contact Email" },
    { key: "pickupContactPhone", header: "Pickup Contact Phone" },
    { key: "bolNumber", header: "BOL Number" },
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
    { key: "customerPo", header: "PO Number: 10 digit 450, 700" },
    { key: "notes", header: "Notes" },
    { key: "sku", header: "SKU" },
    { key: "stackable", header: "Stackable" },
    { key: "expedited", header: "Expedited" },
  ],
};
