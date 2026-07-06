/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: xml.js                                                               ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Wraps browser XML APIs with small helpers used to inspect and edit         ║
 * ║ workbook XML safely.                                                       ║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ Workbook XML errors should become clear user-facing failures rather than   ║
 * ║ silent malformed DOM behavior.                                             ║
 * ║                                                                            ║
 * ║ MAIN CLASSES / FUNCTIONS                                                   ║
 * ║ • parseXml — Workbook XML errors should become clear user-facing           ║
 * ║ failures rather than silent malformed DOM behavior.                        ║
 * ║ • elementChildren — Workbook XML uses namespaces, so callers need a        ║
 * ║ simple way to select direct children by local name.                        ║
 * ║ • firstDescendant — Some workbook XML nodes may be nested, so callers      ║
 * ║ need a safe descendant lookup by local name.                               ║
 * ║ • createElementLike — New worksheet XML nodes must share their parent's    ║
 * ║ namespace to remain valid OOXML.                                           ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import { UserFacingError } from "./errors";

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: parseXml                                                         │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • xmlText and a human-readable label.
 *
 * Returns
 *   • Parsed XML Document or UserFacingError.
 *
 * Why / How
 *   Workbook XML errors should become clear user-facing failures rather than
 *   silent malformed DOM behavior.
 *   It parses the text, checks for parsererror nodes, and throws
 *   UserFacingError with a labeled message when needed.
 */
export function parseXml(xmlText, label) {
  const xmlDocument = new DOMParser().parseFromString(xmlText, "application/xml");
  const parserError = xmlDocument.getElementsByTagName("parsererror")[0];

  if (parserError) {
    throw new UserFacingError(`Could not read ${label} inside the template workbook.`);
  }

  return xmlDocument;
}

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: elementChildren                                                  │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Parent element and desired XML local name.
 *
 * Returns
 *   • Array of matching direct child elements.
 *
 * Why / How
 *   Workbook XML uses namespaces, so callers need a simple way to select
 *   direct children by local name.
 *   It filters the element's direct children by localName.
 */
export function elementChildren(element, localName) {
  return Array.from(element?.children ?? []).filter(
    (child) => child.localName === localName,
  );
}

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: firstDescendant                                                  │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Parent element/document and desired local name.
 *
 * Returns
 *   • Matching element or undefined.
 *
 * Why / How
 *   Some workbook XML nodes may be nested, so callers need a safe descendant
 *   lookup by local name.
 *   It walks all descendants and returns the first matching local name.
 */
export function firstDescendant(element, localName) {
  return Array.from(element?.getElementsByTagName("*") ?? []).find(
    (child) => child.localName === localName,
  );
}

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: createElementLike                                                │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • XML document, parent element, local name.
 *
 * Returns
 *   • New namespaced XML element.
 *
 * Why / How
 *   New worksheet XML nodes must share their parent's namespace to remain
 *   valid OOXML.
 *   It creates an element using the parent namespace URI.
 */
export function createElementLike(xmlDocument, parent, localName) {
  return xmlDocument.createElementNS(parent.namespaceURI, localName);
}
