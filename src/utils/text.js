/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: text.js                                                              ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Contains shared text, numeric, header, and rounding helpers so Excel       ║
 * ║ parsing rules stay consistent.                                             ║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ Excel parsing frequently needs one shared definition of blank across       ║
 * ║ nulls, undefined values, and whitespace-only text.                         ║
 * ║                                                                            ║
 * ║ MAIN CLASSES / FUNCTIONS                                                   ║
 * ║ • isBlank — Excel parsing frequently needs one shared definition of        ║
 * ║ blank across nulls, undefined values, and whitespace-only text.            ║
 * ║ • normalizeText — Case-insensitive comparisons need a consistent           ║
 * ║ normalization rule everywhere in the app.                                  ║
 * ║ • normalizeHeader — Template headers contain presentation wording such     ║
 * ║ as "(Dropdown options)" that should not prevent layout matching.           ║
 * ║ • getCellDisplayValue — Excel cells may expose a formatted display value   ║
 * ║ or an underlying raw value, and header/error text should prefer the        ║
 * ║ display form.                                                              ║
 * ║ • toFiniteNumber — Business calculations must not accidentally perform     ║
 * ║ arithmetic on arbitrary text.                                              ║
 * ║ • numericWhenPossible — Template cells should receive numbers when safe    ║
 * ║ but preserve meaningful nonnumeric values such as error markers.           ║
 * ║ • formatStoreNumber — Older store matching logic needs a safe way to       ║
 * ║ normalize numeric-looking store values without turning 2077 into 20770.    ║
 * ║ • formatNumber — Preview tables should show readable numeric values        ║
 * ║ instead of long floating-point strings.                                    ║
 * ║ • roundToDecimalPlaces — The receiving template accepts at most three      ║
 * ║ decimals for pallet fractions.                                             ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: isBlank                                                          │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • value: any value.
 *
 * Returns
 *   • True when value should be treated as blank.
 *
 * Why / How
 *   Excel parsing frequently needs one shared definition of blank across
 *   nulls, undefined values, and whitespace-only text.
 *   It checks nullish values first and otherwise trims string form.
 */
export function isBlank(value) {
  return value === null || value === undefined || String(value).trim() === "";
}

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: normalizeText                                                    │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • value: any text-like value.
 *
 * Returns
 *   • Normalized comparison string.
 *
 * Why / How
 *   Case-insensitive comparisons need a consistent normalization rule
 *   everywhere in the app.
 *   It normalizes Unicode, converts nonbreaking spaces, collapses
 *   whitespace, trims, and lowercases.
 */
export function normalizeText(value) {
  return String(value ?? "")
    .normalize("NFKC")
    .replace(/\u00a0/g, " ")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: normalizeHeader                                                  │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Raw template header value.
 *
 * Returns
 *   • Normalized header comparison key.
 *
 * Why / How
 *   Template headers contain presentation wording such as "(Dropdown
 *   options)" that should not prevent layout matching.
 *   It normalizes text, removes the dropdown annotation, and converts
 *   punctuation runs to spaces.
 */
export function normalizeHeader(value) {
  return normalizeText(value)
    .replace(/\(\s*dropdown(?:\s+options?)?\s*\)/gi, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: getCellDisplayValue                                              │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • SheetJS cell object.
 *
 * Returns
 *   • String display value.
 *
 * Why / How
 *   Excel cells may expose a formatted display value or an underlying raw
 *   value, and header/error text should prefer the display form.
 *   It returns cell.w when present, otherwise cell.v, otherwise blank.
 */
export function getCellDisplayValue(cell) {
  if (!cell) return "";
  if (cell.w !== undefined && cell.w !== null && cell.w !== "") return String(cell.w);
  if (cell.v === undefined || cell.v === null) return "";
  return String(cell.v);
}

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: toFiniteNumber                                                   │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • value: raw candidate number.
 *
 * Returns
 *   • Finite number or null.
 *
 * Why / How
 *   Business calculations must not accidentally perform arithmetic on
 *   arbitrary text.
 *   It accepts finite numbers directly or removes commas and parses numeric
 *   text, returning null on failure.
 */
export function toFiniteNumber(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }

  if (isBlank(value)) return null;

  const normalized = String(value).replace(/,/g, "").trim();
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: numericWhenPossible                                              │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • value: raw candidate.
 *
 * Returns
 *   • Number when parseable; otherwise original value.
 *
 * Why / How
 *   Template cells should receive numbers when safe but preserve meaningful
 *   nonnumeric values such as error markers.
 *   It tries the finite parser and falls back to the original value.
 */
export function numericWhenPossible(value) {
  const numeric = toFiniteNumber(value);
  return numeric === null ? value : numeric;
}

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: formatStoreNumber                                                │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Raw store value and target length.
 *
 * Returns
 *   • Zero-padded store number or empty string.
 *
 * Why / How
 *   Older store matching logic needs a safe way to normalize numeric-looking
 *   store values without turning 2077 into 20770.
 *   It accepts one numeric source, strips decimal .0 when appropriate, then
 *   pads to requested length.
 */
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

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: formatNumber                                                     │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Number-like value and max digits.
 *
 * Returns
 *   • Formatted numeric string.
 *
 * Why / How
 *   Preview tables should show readable numeric values instead of long
 *   floating-point strings.
 *   It delegates to locale formatting with a maximum fraction-digit limit.
 */
export function formatNumber(value, digits = 2) {
  return Number(value).toLocaleString(undefined, {
    maximumFractionDigits: digits,
  });
}

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: roundToDecimalPlaces                                             │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Numeric value and allowed decimals.
 *
 * Returns
 *   • Rounded number, or original value if not finite.
 *
 * Why / How
 *   The receiving template accepts at most three decimals for pallet
 *   fractions.
 *   It uses a power-of-ten multiplier plus EPSILON to reduce common
 *   floating-point rounding artifacts.
 */
export function roundToDecimalPlaces(value, decimalPlaces = 3) {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return value;
  }

  const multiplier = 10 ** decimalPlaces;

  return Math.round((value + Number.EPSILON) * multiplier) / multiplier;
}
