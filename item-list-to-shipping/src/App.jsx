import { useMemo, useState } from "react";
import { APP_CONFIG, TEMPLATE_FIELDS } from "./config/transformConfig";
import { InputWorkbookReader } from "./services/InputWorkbookReader";
import { TemplateWorkbookService } from "./services/TemplateWorkbookService";
import { downloadBlob } from "./utils/download";
import { formatNumber } from "./utils/text";
import "./App.css";

function App() {
  const services = useMemo(
    () => ({
      inputReader: new InputWorkbookReader(APP_CONFIG.input),
      templateWorkbook: new TemplateWorkbookService({
        appConfig: APP_CONFIG,
        templateFields: TEMPLATE_FIELDS,
      }),
    }),
    [],
  );

  const [sourceFile, setSourceFile] = useState(null);
  const [shipment, setShipment] = useState(null);
  const [error, setError] = useState("");
  const [status, setStatus] = useState("");
  const [isBusy, setIsBusy] = useState(false);

  async function loadSourceFile(file) {
    if (!file) return;

    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      setError("Please choose an .xlsx source workbook.");
      return;
    }

    setIsBusy(true);
    setError("");
    setStatus("Reading input workbook…");

    try {
      const parsedShipment = await services.inputReader.readFile(file);
      setSourceFile(file);
      setShipment(parsedShipment);

      const issueText = parsedShipment.issues.length
        ? ` ${parsedShipment.issues.length} issue(s) were logged in the browser console.`
        : "";

      setStatus(
        `Ready: ${parsedShipment.itemCount} item row(s) across ${parsedShipment.skids.length} skid(s).${issueText}`,
      );
    } catch (caughtError) {
      console.error("[Fuse Order Template Filler]", {
        type: "INPUT_WORKBOOK_FATAL_ERROR",
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

  async function generateTemplate() {
    if (!shipment || !sourceFile) return;

    setIsBusy(true);
    setError("");
    setStatus("Validating dropdowns and building a copy of the template…");

    try {
      const { blob, outputRows, issues } =
        await services.templateWorkbook.createFilledTemplate(shipment);
      downloadBlob(blob, APP_CONFIG.template.outputFileName);
      setStatus(
        `Downloaded ${outputRows.length} populated template row(s). ${issues.length} issue(s) were logged in the browser console.`,
      );
    } catch (caughtError) {
      console.error("[Fuse Order Template Filler]", {
        type: "TEMPLATE_GENERATION_FATAL_ERROR",
        location: APP_CONFIG.template.url,
        message: caughtError.message || "Could not create the filled template.",
        resolution: "Check the template workbook and configuration, then try again.",
      });
      setError(caughtError.message || "Could not create the filled template.");
      setStatus("");
    } finally {
      setIsBusy(false);
    }
  }

  function onDrop(event) {
    event.preventDefault();
    void loadSourceFile(event.dataTransfer.files?.[0]);
  }

  const previewRows = shipment
    ? shipment.skids.flatMap((skid) =>
        skid.items.map((item, itemIndex) => ({
          sourceRow: item.sourceRow,
          type: itemIndex === 0 ? "order" : "commodity",
          customerPo: item.customerPo,
          itemCode: item.oldItemCode,
          quantity: item.quantity,
          targetStore: item.targetStore || "(blank — Destination left blank)",
          dimensions: `${skid.length} × ${skid.width} × ${skid.height}`,
          splitWeight: skid.perItemWeight,
          palletFraction: skid.palletFraction,
        })),
      )
    : [];

  return (
    <main className="appShell">
      <section className="hero">
        <p className="eyebrow">Fuse Order</p>
        <h1>Template Filler</h1>
        <p>
          Upload the source workbook. The app reads merged skid groups, validates the template dropdowns,
          and downloads a filled copy without changing the original template file.
        </p>
      </section>

      <section
        className="dropZone"
        onDragOver={(event) => event.preventDefault()}
        onDrop={onDrop}
      >
        <strong>Drop the source .xlsx workbook here</strong>
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

      {sourceFile && <p className="fileName">Loaded: {sourceFile.name}</p>}
      {status && <p className="status">{status}</p>}
      {error && <p className="error" role="alert">{error}</p>}

      {shipment && (
        <section className="previewSection">
          <div className="previewHeader">
            <div>
              <h2>Input preview</h2>
              <p>Customer PO is copied from source column B for each item row.</p>
            </div>

            <button className="primaryButton" onClick={() => void generateTemplate()} disabled={isBusy}>
              {isBusy ? "Working…" : "Download Filled Template"}
            </button>
          </div>

          <div className="tableWrap">
            <table>
              <thead>
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
              </thead>
              <tbody>
                {previewRows.map((row) => (
                  <tr key={row.sourceRow}>
                    <td>{row.sourceRow}</td>
                    <td>{row.type}</td>
                    <td>{row.customerPo}</td>
                    <td>{row.itemCode}</td>
                    <td>{row.quantity}</td>
                    <td>{row.targetStore}</td>
                    <td>{row.dimensions}</td>
                    <td>{formatPreviewValue(row.splitWeight, 4)}</td>
                    <td>{formatPreviewValue(row.palletFraction, 4)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}
    </main>
  );
}

function formatPreviewValue(value, digits) {
  return typeof value === "number" ? formatNumber(value, digits) : value;
}

export default App;
