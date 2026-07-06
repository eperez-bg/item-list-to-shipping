/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: eslint.config.js                                                     ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Defines the lint rules for browser-based JavaScript and React files.       ║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ Keeping this responsibility in one focused file makes the spreadsheet      ║
 * ║ workflow easier to understand, change, and test without affecting          ║
 * ║ unrelated code.                                                            ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

import js from "@eslint/js";
import globals from "globals";
import reactHooks from "eslint-plugin-react-hooks";
import reactRefresh from "eslint-plugin-react-refresh";
import { defineConfig, globalIgnores } from "eslint/config";

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: unnamed function                                                 │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Static lint configuration.
 *
 * Returns
 *   • ESLint flat configuration array.
 *
 * Why / How
 *   Linting catches common React, hook, and JavaScript mistakes before they
 *   * become browser bugs. The generated dist folder is intentionally
 *   ignored * because it is build output rather than source code.
 *   Combines ESLint recommended rules with React Hooks and Vite refresh
 *   rules.
 */
export default defineConfig([
  globalIgnores(["dist"]),
  {
    files: ["**/*.{js,jsx}"],
    extends: [
      js.configs.recommended,
      reactHooks.configs.flat.recommended,
      reactRefresh.configs.vite,
    ],
    languageOptions: {
      globals: globals.browser,
      parserOptions: { ecmaFeatures: { jsx: true } },
    },
  },
]);
