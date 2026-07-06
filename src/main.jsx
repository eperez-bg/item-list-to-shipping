/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: main.jsx                                                             ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Bootstraps React by mounting the App component into the single root        ║
 * ║ element provided by index.html.                                            ║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ Keeping this responsibility in one focused file makes the spreadsheet      ║
 * ║ workflow easier to understand, change, and test without affecting          ║
 * ║ unrelated code.                                                            ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import App from "./App";
import "./index.css";

createRoot(document.getElementById("root")).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
