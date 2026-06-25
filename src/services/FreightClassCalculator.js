import { toFiniteNumber } from "../utils/text";

const DENSITY_CLASS_BANDS = [
  { minimumDensity: 50, freightClass: 50 },
  { minimumDensity: 35, freightClass: 55 },
  { minimumDensity: 30, freightClass: 60 },
  { minimumDensity: 22.5, freightClass: 65 },
  { minimumDensity: 15, freightClass: 70 },
  { minimumDensity: 13.5, freightClass: 77.5 },
  { minimumDensity: 12, freightClass: 85 },
  { minimumDensity: 10.5, freightClass: 92.5 },
  { minimumDensity: 9, freightClass: 100 },
  { minimumDensity: 8, freightClass: 110 },
  { minimumDensity: 7, freightClass: 125 },
  { minimumDensity: 6, freightClass: 150 },
  { minimumDensity: 5, freightClass: 175 },
  { minimumDensity: 4, freightClass: 200 },
  { minimumDensity: 3, freightClass: 250 },
  { minimumDensity: 2, freightClass: 300 },
  { minimumDensity: 1, freightClass: 400 },
  { minimumDensity: 0, freightClass: 500 },
];

const ALLOWED_FREIGHT_CLASSES = new Set(
  DENSITY_CLASS_BANDS.map((band) => band.freightClass),
);

export class FreightClassCalculator {
  constructor(freightConfig = {}, issueReporter) {
    this.customClassifier =
      typeof freightConfig.classifier === "function"
        ? freightConfig.classifier
        : null;

    this.issueReporter = issueReporter;
  }

  /*
    Freight Class is calculated per output item row. The item receives:
      - the full L/W/H from its physical merged-dimension group
      - its own allocated G.W. share from the source G.W. range

    This intentionally allows two items in the same PO to have different
    classes when they originated from different dimension groups.
  */
  calculate(item) {
    const location = `Input row ${item.sourceRow}`;
    const length = positiveNumberOrNull(item.length);
    const width = positiveNumberOrNull(item.width);
    const height = positiveNumberOrNull(item.height);
    const grossWeight = positiveNumberOrNull(item.allocatedWeight);

    if (!length || !width || !height || !grossWeight) {
      this.issueReporter.record({
        type: "FREIGHT_CLASS_SKIPPED",
        location,
        field: "Freight Class",
        message:
          "Freight Class could not be calculated because this item has one or more missing or invalid dimensions or allocated gross-weight values.",
        resolution:
          "Freight Class is a template dropdown, so it will be left blank.",
      });

      return "";
    }

    const cubicFeet = (length * width * height) / 1728;
    const densityPcf = grossWeight / cubicFeet;

    if (!Number.isFinite(cubicFeet) || cubicFeet <= 0) {
      this.issueReporter.record({
        type: "INVALID_FREIGHT_VOLUME",
        location,
        field: "Freight Class",
        message:
          "The item's volume could not be calculated from Length, Width, and Height.",
        resolution:
          "Freight Class is a template dropdown, so it will be left blank.",
      });

      return "";
    }

    if (!Number.isFinite(densityPcf) || densityPcf <= 0) {
      this.issueReporter.record({
        type: "INVALID_FREIGHT_DENSITY",
        location,
        field: "Freight Class",
        message:
          "The item's freight density could not be calculated from its dimensions and allocated gross weight.",
        resolution:
          "Freight Class is a template dropdown, so it will be left blank.",
      });

      return "";
    }

    const calculatedClass = this.customClassifier
      ? this.customClassifier({
          length,
          width,
          height,
          grossWeight,
          cubicFeet,
          densityPcf,
          item,
        })
      : this.getDensityBasedFreightClass(densityPcf);

    const freightClass = this.normalizeAllowedClass(calculatedClass);

    if (freightClass === null) {
      this.issueReporter.record({
        type: "INVALID_FREIGHT_CLASS_RESULT",
        location,
        field: "Freight Class",
        message:
          `The freight-class calculation returned "${calculatedClass}", ` +
          "which is not one of the allowed template dropdown values.",
        resolution:
          "Freight Class is a template dropdown, so it will be left blank.",
      });

      return "";
    }

    console.info("[Fuse Order Template Filler]", {
      type: "FREIGHT_CLASS_CALCULATED",
      location,
      dimensionSourceRows:
        item.dimensionSourceStartRow && item.dimensionSourceEndRow
          ? `${item.dimensionSourceStartRow}-${item.dimensionSourceEndRow}`
          : null,
      length,
      width,
      height,
      grossWeight,
      cubicFeet: Number(cubicFeet.toFixed(4)),
      densityPcf: Number(densityPcf.toFixed(4)),
      freightClass,
    });

    return freightClass;
  }

  getDensityBasedFreightClass(densityPcf) {
    const matchingBand = DENSITY_CLASS_BANDS.find(
      (band) => densityPcf >= band.minimumDensity,
    );

    return matchingBand?.freightClass ?? 500;
  }

  normalizeAllowedClass(value) {
    const numericValue = Number(value);

    if (!Number.isFinite(numericValue)) {
      return null;
    }

    return ALLOWED_FREIGHT_CLASSES.has(numericValue)
      ? numericValue
      : null;
  }
}

function positiveNumberOrNull(value) {
  const numericValue = toFiniteNumber(value);
  return numericValue !== null && numericValue > 0 ? numericValue : null;
}
