/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: importIssues.js                                                      ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Defines structured recoverable-import issues and logs them in a            ║
 * ║ consistent DevTools-friendly format.                                       ║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ Recoverable spreadsheet problems need a structured record for both UI      ║
 * ║ summaries and developer-console investigation.                             ║
 * ║                                                                            ║
 * ║ MAIN CLASSES / FUNCTIONS                                                   ║
 * ║ • ImportIssue — Recoverable spreadsheet problems need a structured         ║
 * ║ record for both UI summaries and developer-console investigation.          ║
 * ║ • ImportIssueReporter — Multiple services need to log recoverable          ║
 * ║ problems in the same format without each duplicating console code.         ║
 * ║ • record — Each recoverable problem should be collected and visible in     ║
 * ║ DevTools without interrupting the entire import.                           ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

export const ERROR_VALUE = "ERROR";

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ CLASS: ImportIssue                                                         │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Structured issue properties.
 *
 * Returns
 *   • An ImportIssue instance.
 *
 * Why / How
 *   Recoverable spreadsheet problems need a structured record for both UI
 *   summaries and developer-console investigation.
 *   It stores the issue category, location, field, message, and chosen
 *   resolution.
 */
export class ImportIssue {
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ CONSTRUCTOR: constructor                                                   │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • existingIssues: optional array.
   *
   * Returns
   *   • A configured ImportIssueReporter.
   *
   * Why / How
   *   Workbook parsing and template generation should contribute to one
   *   combined issue list.
   *   It stores the provided array by reference so additions remain visible to
   *   callers.
   */
  constructor({ type, location, field, message, resolution }) {
    this.type = type;
    this.location = location;
    this.field = field;
    this.message = message;
    this.resolution = resolution;
  }
}

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ CLASS: ImportIssueReporter                                                 │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • Optional existing issue array.
 *
 * Returns
 *   • A reusable reporter.
 *
 * Why / How
 *   Multiple services need to log recoverable problems in the same format
 *   without each duplicating console code.
 *   It owns the shared issue array and creates structured ImportIssue
 *   records.
 */
export class ImportIssueReporter {
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ CONSTRUCTOR: constructor                                                   │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • existingIssues: optional array.
   *
   * Returns
   *   • A configured ImportIssueReporter.
   *
   * Why / How
   *   Workbook parsing and template generation should contribute to one
   *   combined issue list.
   *   It stores the provided array by reference so additions remain visible to
   *   callers.
   */
  constructor(existingIssues = []) {
    this.issues = existingIssues;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: record                                                             │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • Structured issue metadata.
   *
   * Returns
   *   • The newly created ImportIssue.
   *
   * Why / How
   *   Each recoverable problem should be collected and visible in DevTools
   *   without interrupting the entire import.
   *   It creates an ImportIssue, appends it, logs a structured console object,
   *   and returns the issue.
   */
  record({ type, location, field, message, resolution }) {
    const issue = new ImportIssue({
      type,
      location,
      field,
      message,
      resolution,
    });

    this.issues.push(issue);

    // This is intentionally a structured console message so it is easy to
    // filter by type or inspect the exact source/template location in DevTools.
    console.error("[Fuse Order Template Filler]", {
      type: issue.type,
      location: issue.location,
      field: issue.field,
      message: issue.message,
      resolution: issue.resolution,
    });

    return issue;
  }
}
