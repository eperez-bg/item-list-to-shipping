/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: download.js                                                          ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Starts a browser download from an in-memory Blob without sending the       ║
 * ║ generated workbook to a server.                                            ║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ The finished workbook should download locally without uploading user       ║
 * ║ data to a server.                                                          ║
 * ║                                                                            ║
 * ║ MAIN CLASSES / FUNCTIONS                                                   ║
 * ║ • downloadBlob — The finished workbook should download locally without     ║
 * ║ uploading user data to a server.                                           ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: downloadBlob                                                     │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • blob: generated workbook bytes; filename: download name.
 *
 * Returns
 *   • No returned value; prompts a browser download.
 *
 * Why / How
 *   The finished workbook should download locally without uploading user
 *   data to a server.
 *   It creates a temporary object URL and anchor, clicks it
 *   programmatically, then revokes the URL after the browser starts the
 *   download.
 */
export function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();

  window.setTimeout(() => URL.revokeObjectURL(url), 1_000);
}
