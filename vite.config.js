/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: vite.config.js                                                       ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Configures Vite to build the React app and emit correct asset paths *      ║
 * ║ for the GitHub Pages repository subfolder.                                 ║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ Keeping this responsibility in one focused file makes the spreadsheet      ║
 * ║ workflow easier to understand, change, and test without affecting          ║
 * ║ unrelated code.                                                            ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: unnamed function                                                 │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Static build configuration.
 *
 * Returns
 *   • Vite configuration object.
 *
 * Why / How
 *   GitHub Pages serves this project below /item-list-to-shipping/, not at *
 *   the domain root. Without base, built JavaScript and template URLs would
 *   use * incorrect root-relative paths after deployment.
 *   Vite prefixes produced asset URLs with this base during npm run build.
 */
export default defineConfig({
  plugins: [react()],
  base: "/item-list-to-shipping/",
});
