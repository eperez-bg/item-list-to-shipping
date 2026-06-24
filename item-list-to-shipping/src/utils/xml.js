import { UserFacingError } from "./errors";

export function parseXml(xmlText, label) {
  const xmlDocument = new DOMParser().parseFromString(xmlText, "application/xml");
  const parserError = xmlDocument.getElementsByTagName("parsererror")[0];

  if (parserError) {
    throw new UserFacingError(`Could not read ${label} inside the template workbook.`);
  }

  return xmlDocument;
}

export function elementChildren(element, localName) {
  return Array.from(element?.children ?? []).filter(
    (child) => child.localName === localName,
  );
}

export function firstDescendant(element, localName) {
  return Array.from(element?.getElementsByTagName("*") ?? []).find(
    (child) => child.localName === localName,
  );
}

export function createElementLike(xmlDocument, parent, localName) {
  return xmlDocument.createElementNS(parent.namespaceURI, localName);
}
