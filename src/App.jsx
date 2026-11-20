/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: App.jsx                                                              ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Defines the React screen and coordinates file upload, the LTL/FTL mode     ║
 * ║ toggle, preview data, and template downloads.                              ║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ It is the composition root for the visible application, so UI state and    ║
 * ║ service creation stay in one place.                                        ║
 * ║                                                                            ║
 * ║ MAIN CLASSES / FUNCTIONS                                                   ║
 * ║ • App — It is the composition root for the visible application, so UI      ║
 * ║ state and service creation stay in one place.                              ║
 * ║ • selectShippingMode — Changing modes changes both parsing rules and       ║
 * ║ output templates, so old parsed data must not be reused accidentally.      ║
 * ║ • loadSourceFile — The UI needs one controlled path for validating and     ║
 * ║ parsing uploads so all errors are shown consistently.                      ║
 * ║ • generateTemplate — Generating a workbook touches several services and    ║
 * ║ can fail, so it is kept behind one button handler with status updates.     ║
 * ║ • onDrop — Drag-and-drop needs to prevent the browser's default            ║
 * ║ file-opening behavior before using the dropped file.                       ║
 * ║ • buildPreviewRows — The preview should mirror the ordering and            ║
 * ║ calculations used by the downloaded workbook so users can spot issues      ║
 * ║ early.                                                                     ║
 * ║ • createSkidIndex — LTL preview generation needs quick access from each    ║
 * ║ item to its physical merged-dimension group.                               ║
 * ║ • previewDescription — The UI should explain the grouping rule that the    ║
 * ║ currently selected mode will apply.                                        ║
 * ║ • formatPreviewValue — Raw floating-point values are difficult to scan,    ║
 * ║ especially for weights and fractional pallet shares.                       ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import { useMemo, useState } from "react";
import {
  APP_CONFIG,
  SHIPPING_MODES,
  TEMPLATE_FIELDS_BY_MODE,
} from "./config/transformConfig";
import { InputWorkbookReader } from "./services/InputWorkbookReader";
import { TemplateWorkbookService } from "./services/TemplateWorkbookService";
import { downloadBlob } from "./utils/download";
import { ERROR_VALUE } from "./utils/importIssues";
import { formatNumber, roundToDecimalPlaces, toFiniteNumber } from "./utils/text";
import "./App.css";

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: App                                                              │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • None; React calls the component.
 *
 * Returns
 *   • React JSX for the full application screen.
 *
 * Why / How
 *   It is the composition root for the visible application, so UI state and
 *   service creation stay in one place.
 *   It stores the active mode, file, parsed shipment, status, and busy
 *   state; then renders the appropriate controls and preview.
 */
function App() {
  const [shippingMode, setShippingMode] = useState(SHIPPING_MODES.LTL);
  const [sourceFile, setSourceFile] = useState(null);
  const [shipment, setShipment] = useState(null);
  const [error, setError] = useState("");
  const [status, setStatus] = useState("");
  const [isBusy, setIsBusy] = useState(false);

  const services = useMemo(() => {
    const templateConfig = APP_CONFIG.templates[shippingMode];

    return {
      inputReader: new InputWorkbookReader(APP_CONFIG.input, shippingMode),
      templateWorkbook: new TemplateWorkbookService({
        appConfig: APP_CONFIG,
        templateConfig,
        templateFields: TEMPLATE_FIELDS_BY_MODE[shippingMode],
        shippingMode,
      }),
    };
  }, [shippingMode]);

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ FUNCTION: selectShippingMode                                               │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • nextMode: the requested LTL or FTL mode.
   *
   * Returns
   *   • No returned value; React state is reset for the new mode.
   *
   * Why / How
   *   Changing modes changes both parsing rules and output templates, so old
   *   parsed data must not be reused accidentally.
   *   It updates the mode and clears the selected file, shipment, status, and
   *   error state.
   */
  function selectShippingMode(nextMode) {
    if (nextMode === shippingMode) return;

    // Parsing and output rules differ by mode, so remove the prior result.
    setShippingMode(nextMode);
    setSourceFile(null);
    setShipment(null);
    setError("");
    setStatus("");
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ FUNCTION: loadSourceFile                                                   │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • file: a browser File selected or dropped by the user.
   *
   * Returns
   *   • A Promise that resolves after React state reflects the parsing result.
   *
   * Why / How
   *   The UI needs one controlled path for validating and parsing uploads so
   *   all errors are shown consistently.
   *   It validates the extension, calls the active InputWorkbookReader, then
   *   stores either the parsed shipment or a user-facing error.
   */
  async function loadSourceFile(file) {
    if (!file) return;

    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      setError("Please choose an .xlsx source workbook.");
      return;
    }

    setIsBusy(true);
    setError("");
    setStatus(`Reading ${shippingMode} input workbook…`);

    try {
      const parsedShipment = await services.inputReader.readFile(file);
      setSourceFile(file);
      setShipment(parsedShipment);

      const issueText = parsedShipment.issues.length
        ? ` ${parsedShipment.issues.length} issue(s) were logged in the browser console.`
        : "";
      const groupingText = shippingMode === SHIPPING_MODES.FTL
        ? `${parsedShipment.truckCount} Truck #(s)`
        : `${parsedShipment.storeCount} Store #(s)`;

      setStatus(
        `Ready: ${parsedShipment.itemCount} item row(s) across ${groupingText}.${issueText}`,
      );
    } catch (caughtError) {
      console.error("[Fuse Order Template Filler]", {
        type: "INPUT_WORKBOOK_FATAL_ERROR",
        shippingMode,
        location: file.name,
        message: caughtError.message || "Could not read the input workbook.",
        resolution: "Fix the workbook or configuration, then upload it again.",
      });
      setSourceFile(null);
      setShipment(null);
      setStatus("");
      setError(caughtError.message || "Could not read the input workbook.");
    } finally {
      setIsBusy(false);
    }
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ FUNCTION: generateTemplate                                                 │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None; uses the currently parsed shipment and active mode.
   *
   * Returns
   *   • A Promise that resolves after the download begins or an error is
   *   • displayed.
   *
   * Why / How
   *   Generating a workbook touches several services and can fail, so it is
   *   kept behind one button handler with status updates.
   *   It asks the active TemplateWorkbookService to create a Blob, then
   *   delegates the browser download to downloadBlob.
   */
  async function generateTemplate() {
    if (!shipment || !sourceFile) return;

    setIsBusy(true);
    setError("");
    setStatus(`Validating dropdowns and building a ${shippingMode} template copy…`);

    try {
      const { blob, outputRows, issues } =
        await services.templateWorkbook.createFilledTemplate(shipment);

      downloadBlob(blob, services.templateWorkbook.outputFileName);
      setStatus(
        `Downloaded ${outputRows.length} populated ${shippingMode} template row(s). ${issues.length} issue(s) were logged in the browser console.`,
      );
    } catch (caughtError) {
      console.error("[Fuse Order Template Filler]", {
        type: "TEMPLATE_GENERATION_FATAL_ERROR",
        shippingMode,
        location: services.templateWorkbook.templateUrl,
        message: caughtError.message || "Could not create the filled template.",
        resolution: "Check the template workbook and configuration, then try again.",
      });
      setError(caughtError.message || "Could not create the filled template.");
      setStatus("");
    } finally {
      setIsBusy(false);
    }
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ FUNCTION: onDrop                                                           │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • event: a React drag event.
   *
   * Returns
   *   • No returned value; starts asynchronous file loading.
   *
   * Why / How
   *   Drag-and-drop needs to prevent the browser's default file-opening
   *   behavior before using the dropped file.
   *   It cancels the default event and passes the first dropped file into the
   *   same upload flow used by the file picker.
   */
  function onDrop(event) {
    event.preventDefault();
    void loadSourceFile(event.dataTransfer.files?.[0]);
  }

  const previewRows = shipment
    ? buildPreviewRows(shipment, shippingMode)
    : [];

  return (
    <main className="appShell">
      <section className="hero">
        <p className="eyebrow">Fuse Order</p>
        <h1>Template Filler</h1>
        <p>
          Upload the source workbook. The app reads Excel data in your browser,
          validates template dropdowns, and downloads a filled copy without
          changing either original template file.
        </p>
      </section>

      <section className="shippingModeSection" aria-label="Shipping mode selection">
        <span className="shippingModeLabel">Shipping Mode</span>
        <div className="shippingModeToggle" role="radiogroup" aria-label="Shipping mode">
          <button
            type="button"
            role="radio"
            aria-checked={shippingMode === SHIPPING_MODES.LTL}
            className={`shippingModeButton ${
              shippingMode === SHIPPING_MODES.LTL ? "shippingModeButtonActive" : ""
            }`}
            onClick={() => selectShippingMode(SHIPPING_MODES.LTL)}
            disabled={isBusy}
          >
            LTL
          </button>
          <button
            type="button"
            role="radio"
            aria-checked={shippingMode === SHIPPING_MODES.FTL}
            className={`shippingModeButton ${
              shippingMode === SHIPPING_MODES.FTL ? "shippingModeButtonActive" : ""
            }`}
            onClick={() => selectShippingMode(SHIPPING_MODES.FTL)}
            disabled={isBusy}
          >
            FTL
          </button>
        </div>
      </section>

      <section
        className="dropZone"
        onDragOver={(event) => event.preventDefault()}
        onDrop={onDrop}
      >
        <strong>Drop the {shippingMode} source .xlsx workbook here</strong>
        <span>or</span>
        <label className="secondaryButton">
          Choose source workbook
          <input
            type="file"
            accept=".xlsx"
            hidden
            onChange={(event) => void loadSourceFile(event.target.files?.[0])}
          />
        </label>
      </section>

      {shippingMode === SHIPPING_MODES.FTL && (
        <p className="modeHint">
          FTL requires <code>Truck #</code>, <code>Item Code</code>, <code>IF</code>,
          <code>target order casepack qty</code>, and <code>casepack price</code>.
          Repeated Item Codes on one truck
          are combined. Headers are case-insensitive. The FTL catalog must exist at
          <code> public/catalogs/BEELINEGROUPFX_L3.xlsx</code>. A row marked
          <code> additional pallet</code> may omit regular item details.
        </p>
      )}

      {sourceFile && <p className="fileName">Loaded: {sourceFile.name}</p>}
      {status && <p className="status">{status}</p>}
      {error && <p className="error" role="alert">{error}</p>}

      {shipment && (
        <section className="previewSection">
          <div className="previewHeader">
            <div>
              <h2>Input preview</h2>
              <p>{previewDescription(shippingMode)}</p>
            </div>
            <button
              className="primaryButton"
              onClick={() => void generateTemplate()}
              disabled={isBusy}
            >
              {isBusy ? "Working…" : `Download Filled ${shippingMode} Template`}
            </button>
          </div>

          <div className="tableWrap">
            <table>
              <thead>
                {shippingMode === SHIPPING_MODES.FTL ? (
                  <tr>
                    <th>Source Rows</th>
                    <th>Type</th>
                    <th>Truck #</th>
                    <th>Customer PO</th>
                    <th>Item Code</th>
                    <th>Store</th>
                    <th>Pickup IF</th>
                    <th>Casepack Qty</th>
                    <th>Total Weight</th>
                    <th>Total Value</th>
                    <th>Pallets</th>
                  </tr>
                ) : (
                  <tr>
                    <th>Source Row</th>
                    <th>Type</th>
                    <th>Customer PO</th>
                    <th>Old Item Code</th>
                    <th>Qty</th>
                    <th>Store</th>
                    <th>L × W × H</th>
                    <th>Weight / Item</th>
                    <th>Pallet Share</th>
                  </tr>
                )}
              </thead>
              <tbody>
                {previewRows.map((row) =>
                  shippingMode === SHIPPING_MODES.FTL ? (
                    <tr key={row.sourceRow}>
                      <td>{row.sourceRow}</td>
                      <td>{row.type}</td>
                      <td>{row.truckNumber}</td>
                      <td>{row.customerPo}</td>
                      <td>{row.itemCode}</td>
                      <td>{row.targetStore}</td>
                      <td>{row.pickupNumber}</td>
                      <td>{row.quantity}</td>
                      <td>{row.totalWeight}</td>
                      <td>{row.totalValue}</td>
                      <td>{row.pallets}</td>
                    </tr>
                  ) : (
                    <tr key={row.sourceRow}>
                      <td>{row.sourceRow}</td>
                      <td>{row.type}</td>
                      <td>{row.customerPo}</td>
                      <td>{row.itemCode}</td>
                      <td>{row.quantity}</td>
                      <td>{row.targetStore}</td>
                      <td>{row.dimensions}</td>
                      <td>{formatPreviewValue(row.allocatedWeight, 4)}</td>
                      <td>{formatPreviewValue(row.palletFraction, 3)}</td>
                    </tr>
                  ),
                )}
              </tbody>
            </table>
          </div>
        </section>
      )}
    </main>
  );
}

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: buildPreviewRows                                                 │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • shipment: parsed InputShipment; shippingMode: LTL or FTL.
 *
 * Returns
 *   • An array of simple objects used by the preview table.
 *
 * Why / How
 *   The preview should mirror the ordering and calculations used by the
 *   downloaded workbook so users can spot issues early.
 *   It chooses truck groups for FTL or Store/Truck groups plus a skid index for LTL,
 *   then produces display-ready row objects.
 */
function buildPreviewRows(shipment, shippingMode) {
  if (shippingMode === SHIPPING_MODES.FTL) {
    // Preview exactly the same pre-aggregation used by the FTL output factory.
    // Catalog fields are resolved only on download and are not shown here.
    return shipment.trucks.flatMap((truck) =>
      truck.getCombinedItemGroups().map((group) => {
        const item = group.representativeItem;
        return {
          sourceRow: group.sourceRows.join(", "),
          type: group.isOrder ? "order" : "commodity",
          truckNumber: truck.truckNumber,
          customerPo: item.customerPo,
          itemCode: item.itemCode,
          targetStore: item.targetStore || "(blank — Destination left blank)",
          pickupNumber: item.pickupNumber,
          quantity: group.sumQuantity() ?? ERROR_VALUE,
          totalWeight: group.sumWeight() ?? ERROR_VALUE,
          totalValue: group.calculateTotalValue(group.sumQuantity()) ?? ERROR_VALUE,
          pallets: group.palletCount,
        };
      }),
    );
  }

  const skidByItem = createSkidIndex(shipment);

  return shipment.storeTrucks.flatMap((truck) =>
    truck.getCombinedItemGroups().map((combinedGroup) => {
      const item = combinedGroup.representativeItem;
      const metrics = buildCombinedPreviewMetrics(combinedGroup, skidByItem);

      return {
        sourceRow: combinedGroup.sourceRows,
        type: truck.isOrderGroup(combinedGroup) ? "order" : "commodity",
        customerPo: item.customerPo,
        itemCode: item.oldItemCode,
        quantity: metrics.quantity,
        targetStore: item.targetStore || "(blank — Destination left blank)",
        dimensions: `${metrics.length} × ${metrics.width} × ${metrics.height}`,
        allocatedWeight: metrics.totalWeight,
        palletFraction: metrics.palletShare,
      };
    }),
  );
}

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: buildCombinedPreviewMetrics                                    │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • combinedGroup: StoreTruckItemGroup used for one LTL output row.
 *   • skidByItem: Map<InputItem, InputSkid> for pallet share lookup.
 *
 * Returns
 *   • Preview-friendly aggregate quantity, dimensions, weight, and pallet data.
 *
 * Why / How
 *   The LTL preview should reflect the same duplicate-item combining rule as
 *   the downloaded template. It sums calculated carton quantities, allocated
 *   weights, and pallet shares while keeping the largest L/W/H values.
 */
function buildCombinedPreviewMetrics(combinedGroup, skidByItem) {
  return {
    quantity: sumPreviewNumbers(
      combinedGroup.items.map((item) => {
        const rawQuantity = toFiniteNumber(item.quantity);
        const casepackQty = toFiniteNumber(item.casepackQty);

        return rawQuantity !== null && casepackQty !== null && casepackQty > 0
          ? rawQuantity / casepackQty
          : ERROR_VALUE;
      }),
      3,
    ),
    length: maxPreviewNumber(combinedGroup.items.map((item) => item.length)),
    width: maxPreviewNumber(combinedGroup.items.map((item) => item.width)),
    height: maxPreviewNumber(combinedGroup.items.map((item) => item.height)),
    totalWeight: sumPreviewNumbers(
      combinedGroup.items.map((item) => item.allocatedWeight),
      4,
    ),
    palletShare: sumPreviewNumbers(
      combinedGroup.items.map((item) => {
        const skid = skidByItem.get(item);
        return skid ? skid.getPalletShareForItem(item, 3) : ERROR_VALUE;
      }),
      3,
    ),
  };
}

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: sumPreviewNumbers                                              │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • values: values to sum; digits: final decimal precision.
 *
 * Returns
 *   • Rounded numeric sum, or ERROR_VALUE when any value is invalid.
 *
 * Why / How
 *   Preview rows should not hide calculation problems. If any part of a
 *   combined row is invalid, the preview shows the same error text pattern the
 *   generated workbook will use.
 */
function sumPreviewNumbers(values, digits) {
  const numericValues = values.map((value) => toFiniteNumber(value));

  if (numericValues.some((value) => value === null)) {
    return ERROR_VALUE;
  }

  return roundToDecimalPlaces(
    numericValues.reduce((sum, value) => sum + value, 0),
    digits,
  );
}

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: maxPreviewNumber                                               │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • values: candidate dimension values.
 *
 * Returns
 *   • Largest valid numeric value, or ERROR_VALUE when none exist.
 *
 * Why / How
 *   Combined duplicate rows keep the largest Length, Width, and Height. This
 *   helper mirrors that rule for the on-screen preview.
 */
function maxPreviewNumber(values) {
  const numericValues = values
    .map((value) => toFiniteNumber(value))
    .filter((value) => value !== null);

  return numericValues.length > 0 ? Math.max(...numericValues) : ERROR_VALUE;
}

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: createSkidIndex                                                  │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • shipment: parsed InputShipment.
 *
 * Returns
 *   • Map<InputItem, InputSkid>.
 *
 * Why / How
 *   LTL preview generation needs quick access from each item to its physical
 *   merged-dimension group.
 *   It walks every skid once and stores each item as a key in a Map pointing
 *   to its skid.
 */
function createSkidIndex(shipment) {
  const skidByItem = new Map();

  shipment.skids.forEach((skid) => {
    skid.items.forEach((item) => {
      skidByItem.set(item, skid);
    });
  });

  return skidByItem;
}

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: previewDescription                                               │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • shippingMode: LTL or FTL.
 *
 * Returns
 *   • A short explanatory string.
 *
 * Why / How
 *   The UI should explain the grouping rule that the currently selected mode
 *   will apply.
 *   It returns one fixed description for LTL and another for FTL.
 */
function previewDescription(shippingMode) {
  return shippingMode === SHIPPING_MODES.FTL
    ? "FTL items are combined by Truck # and Item Code before quantity, total weight, pallets, and freight are calculated. The first combined item is Order; other items in the truck are Commodity."
    : "Preview and download are grouped by Store #: each Order row is followed by its Commodity rows. Matching Item Code + CustomerPO rows inside the same Store group are combined before output.";
}

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: formatPreviewValue                                               │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • value: preview value; digits: maximum decimals.
 *
 * Returns
 *   • A display-friendly string or original value.
 *
 * Why / How
 *   Raw floating-point values are difficult to scan, especially for weights
 *   and fractional pallet shares.
 *   It formats numbers with the shared formatter and leaves non-numeric
 *   values untouched.
 */
function formatPreviewValue(value, digits) {
  return typeof value === "number" ? formatNumber(value, digits) : value;
}

export default App;
