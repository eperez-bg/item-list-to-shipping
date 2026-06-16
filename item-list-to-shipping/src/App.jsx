import { useState } from "react";
import * as XLSX from "xlsx";
import "./App.css";

const INPUT_HEADER_ROW_INDEX = 0;

// Temporary output headers.
// Later, you can replace these with the yellow template headers.
const OUTPUT_HEADERS = [
  "Item",
  "Description",
  "Length",
  "Width",
  "Height",
  "Quantity",
];

function App() {
  const [rows, setRows] = useState([]);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");

  async function handleFile(file) {
    setError("");
    setRows([]);
    setFileName(file.name);

    try {
      const buffer = await file.arrayBuffer();

      const workbook = XLSX.read(buffer, {
        type: "array",
        cellDates: true,
      });

      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      // Important for merged cells
      fillMergedCells(worksheet);

      const data = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: "",
        raw: false,
      });

      const parsedRows = rowsFromHeaderSheet(data, INPUT_HEADER_ROW_INDEX);

      setRows(parsedRows);
    } catch (err) {
      console.error(err);
      setError("Could not read that Excel file.");
    }
  }

  function downloadOutputWorkbook() {
    if (rows.length === 0) {
      setError("Upload an input file first.");
      return;
    }

    // Temporary simple mapping.
    // Later, this is where you will map blue/green headers to yellow headers.
    const outputRows = rows.map((row) => ({
      Item: row["Item"] || row["ITEM"] || row["Item #"] || "",
      Description: row["Description"] || row["DESC"] || "",
      Length: row["L"] || row["Length"] || "",
      Width: row["W"] || row["Width"] || "",
      Height: row["H"] || row["Height"] || "",
      Quantity: row["Qty"] || row["QTY"] || row["Quantity"] || "",
    }));

    const outputSheet = XLSX.utils.json_to_sheet(outputRows, {
      header: OUTPUT_HEADERS,
    });

    const outputWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(outputWorkbook, outputSheet, "Output");

    XLSX.writeFile(outputWorkbook, "filled-template.xlsx");
  }

  function handleDrop(event) {
    event.preventDefault();

    const file = event.dataTransfer.files?.[0];

    if (!file) return;

    if (!file.name.endsWith(".xlsx")) {
      setError("Please upload an .xlsx file.");
      return;
    }

    handleFile(file);
  }

  function handleDragOver(event) {
    event.preventDefault();
  }

  function handleInputChange(event) {
    const file = event.target.files?.[0];

    if (!file) return;

    if (!file.name.endsWith(".xlsx")) {
      setError("Please upload an .xlsx file.");
      return;
    }

    handleFile(file);
  }

  return (
    <main className="page">
      <h1>Excel Template Filler</h1>

      <div
        className="dropzone"
        onDrop={handleDrop}
        onDragOver={handleDragOver}
      >
        <p>Drop input XLSX file here</p>
        <p>or</p>

        <label className="fileButton">
          Choose File
          <input
            type="file"
            accept=".xlsx"
            onChange={handleInputChange}
            hidden
          />
        </label>
      </div>

      {fileName && <p className="muted">Loaded: {fileName}</p>}

      {error && <p className="error">{error}</p>}

      {rows.length > 0 && (
        <>
          <h2>Preview</h2>
          <p className="muted">Rows found: {rows.length}</p>

          <div className="tableWrap">
            <table>
              <thead>
                <tr>
                  {Object.keys(rows[0]).map((header) => (
                    <th key={header}>{header}</th>
                  ))}
                </tr>
              </thead>

              <tbody>
                {rows.slice(0, 10).map((row, rowIndex) => (
                  <tr key={rowIndex}>
                    {Object.keys(rows[0]).map((header) => (
                      <td key={header}>{row[header]}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <button className="downloadButton" onClick={downloadOutputWorkbook}>
            Download Filled Template
          </button>
        </>
      )}
    </main>
  );
}

function rowsFromHeaderSheet(data, headerRowIndex) {
  const headers = data[headerRowIndex].map((header) =>
    String(header).trim()
  );

  const bodyRows = data.slice(headerRowIndex + 1);

  return bodyRows
    .filter((row) => row.some((cell) => String(cell).trim() !== ""))
    .map((row) => {
      const obj = {};

      headers.forEach((header, index) => {
        if (!header) return;
        obj[header] = row[index] ?? "";
      });

      return obj;
    });
}

function fillMergedCells(worksheet) {
  const merges = worksheet["!merges"] || [];

  merges.forEach((merge) => {
    const topLeftAddress = XLSX.utils.encode_cell(merge.s);
    const topLeftCell = worksheet[topLeftAddress];

    if (!topLeftCell) return;

    for (let row = merge.s.r; row <= merge.e.r; row++) {
      for (let col = merge.s.c; col <= merge.e.c; col++) {
        const address = XLSX.utils.encode_cell({ r: row, c: col });

        if (!worksheet[address]) {
          worksheet[address] = { ...topLeftCell };
        }
      }
    }
  });
}

export default App;