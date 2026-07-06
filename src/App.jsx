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
import { formatNumber } from "./utils/text";
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
        : `${parsedShipment.purchaseOrderCount} Customer PO(s)`;

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
          FTL requires <code>Truck #</code>, <code>Pallet</code>, and <code>China PO</code>.
          Header capitalization does not matter. A row whose <code>container</code> value is
          <code>additional pallet</code> can omit regular item details.
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
                    <th>Source Row</th>
                    <th>Type</th>
                    <th>Truck #</th>
                    <th>Customer PO</th>
                    <th>Old Item Code</th>
                    <th>Store</th>
                    <th>China PO</th>
                    <th>Order Weight</th>
                    <th>Truck Pallets</th>
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
                      <td>{row.chinaPo}</td>
                      <td>{row.orderWeight}</td>
                      <td>{row.truckPallets}</td>
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
 *   It chooses truck groups for FTL or PO groups plus a skid index for LTL,
 *   then produces display-ready row objects.
 */
function buildPreviewRows(shipment, shippingMode) {
  if (shippingMode === SHIPPING_MODES.FTL) {
    return shipment.trucks.flatMap((truck) =>
      truck.items.map((item) => {
        const isOrderRow = truck.isFirstItem(item);

        return {
          sourceRow: item.sourceRow,
          type: isOrderRow ? "order" : "commodity",
          truckNumber: item.truckNumber,
          customerPo: item.customerPo,
          itemCode: item.oldItemCode,
          targetStore: item.targetStore || "(blank — Destination left blank)",
          chinaPo: item.chinaPo,
          orderWeight: isOrderRow ? item.grossWeight : 1,
          truckPallets: isOrderRow ? truck.getGreatestPalletCount() : "",
        };
      }),
    );
  }

  const skidByItem = createSkidIndex(shipment);

  return shipment.purchaseOrders.flatMap((purchaseOrder) =>
    purchaseOrder.items.map((item) => {
      const skid = skidByItem.get(item);

      return {
        sourceRow: item.sourceRow,
        type: purchaseOrder.isFirstItem(item) ? "order" : "commodity",
        customerPo: item.customerPo,
        itemCode: item.oldItemCode,
        quantity: item.quantity,
        targetStore: item.targetStore || "(blank — Destination left blank)",
        dimensions: `${item.length} × ${item.width} × ${item.height}`,
        allocatedWeight: item.allocatedWeight,
        palletFraction: skid ? skid.getPalletShareForItem(item, 3) : "",
      };
    }),
  );
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
    ? "Preview and download are grouped by Truck #: each Order row is followed immediately by every Commodity row in that truck."
    : "Preview and download are grouped by Customer PO: each Order row is followed immediately by every Commodity row for that PO. Pallet shares still use merged-dimension groups.";
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
