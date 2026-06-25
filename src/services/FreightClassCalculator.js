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

  calculate(skid) {
    const rowRange = `${skid.sourceStartRow}-${skid.sourceEndRow}`;

    if (!skid.hasValidFreightMeasurements) {
      this.issueReporter.record({
        type: "FREIGHT_CLASS_SKIPPED",
        location: `Input skid rows ${rowRange}`,
        field: "Freight Class",
        message:
          "Freight Class could not be calculated because one or more source dimensions or gross-weight values are missing or invalid.",
        resolution:
          "Freight Class is a template dropdown, so it will be left blank.",
      });

      return "";
    }

    const length = skid.numericLength;
    const width = skid.numericWidth;
    const height = skid.numericHeight;
    const grossWeight = skid.numericGrossWeight;

    const cubicFeet = (length * width * height) / 1728;
    const densityPcf = grossWeight / cubicFeet;

    if (!Number.isFinite(cubicFeet) || cubicFeet <= 0) {
      this.issueReporter.record({
        type: "INVALID_FREIGHT_VOLUME",
        location: `Input skid rows ${rowRange}`,
        field: "Freight Class",
        message:
          "The skid volume could not be calculated from Length, Width, and Height.",
        resolution:
          "Freight Class is a template dropdown, so it will be left blank.",
      });

      return "";
    }

    if (!Number.isFinite(densityPcf) || densityPcf <= 0) {
      this.issueReporter.record({
        type: "INVALID_FREIGHT_DENSITY",
        location: `Input skid rows ${rowRange}`,
        field: "Freight Class",
        message:
          "Freight density could not be calculated from the skid dimensions and gross weight.",
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
        })
      : this.getDensityBasedFreightClass(densityPcf);

    const freightClass = this.normalizeAllowedClass(calculatedClass);

    if (freightClass === null) {
      this.issueReporter.record({
        type: "INVALID_FREIGHT_CLASS_RESULT",
        location: `Input skid rows ${rowRange}`,
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
      location: `Input skid rows ${rowRange}`,
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
