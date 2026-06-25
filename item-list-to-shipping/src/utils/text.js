export function isBlank(value) {
  return value === null || value === undefined || String(value).trim() === "";
}

export function normalizeText(value) {
  return String(value ?? "")
    .normalize("NFKC")
    .replace(/\u00a0/g, " ")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

export function normalizeHeader(value) {
  return normalizeText(value)
    .replace(/\(\s*dropdown(?:\s+options?)?\s*\)/gi, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

export function getCellDisplayValue(cell) {
  if (!cell) return "";
  if (cell.w !== undefined && cell.w !== null && cell.w !== "") return String(cell.w);
  if (cell.v === undefined || cell.v === null) return "";
  return String(cell.v);
}

export function toFiniteNumber(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }

  if (isBlank(value)) return null;

  const normalized = String(value).replace(/,/g, "").trim();
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

export function numericWhenPossible(value) {
  const numeric = toFiniteNumber(value);
  return numeric === null ? value : numeric;
}

export function formatStoreNumber(value, length) {
  if (isBlank(value)) return "";

  // Excel can expose a source value such as 2077 as a number, or as text such
  // as "2077.0". Both must resolve to the same store number, not "20770".
  if (typeof value === "number") {
    if (!Number.isSafeInteger(value) || value < 0) return "";
    return String(value).padStart(length, "0");
  }

  const text = String(value)
    .normalize("NFKC")
    .replace(/\u00a0/g, " ")
    .replace(/,/g, "")
    .trim();

  const wholeNumber = text.match(/^0*(\d+)(?:\.0+)?$/);
  let digits = wholeNumber?.[1] ?? "";

  if (!digits) {
    const digitGroups = text.match(/\d+/g) ?? [];

    // A label like "Store 2077" is okay. Multiple unrelated numbers are not
    // safe to guess from, so leave that destination blank instead.
    if (digitGroups.length !== 1) return "";
    digits = digitGroups[0];
  }

  return digits.padStart(length, "0");
}

export function formatNumber(value, digits = 2) {
  return Number(value).toLocaleString(undefined, {
    maximumFractionDigits: digits,
  });
}

export function roundToDecimalPlaces(value, decimalPlaces = 3) {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return value;
  }

  const multiplier = 10 ** decimalPlaces;

  return Math.round((value + Number.EPSILON) * multiplier) / multiplier;
}
