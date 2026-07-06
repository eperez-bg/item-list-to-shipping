/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: OutputLoadRow.js                                                     ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Provides a small wrapper around a completed output row so                  ║
 * ║ template-writing code receives a predictable values object.                ║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ Keeping rows behind a tiny model gives the XML patcher one predictable     ║
 * ║ interface instead of depending on factory internals.                       ║
 * ║                                                                            ║
 * ║ MAIN CLASSES / FUNCTIONS                                                   ║
 * ║ • OutputLoadRow — Keeping rows behind a tiny model gives the XML patcher   ║
 * ║ one predictable interface instead of depending on factory internals.       ║
 * ║ • toTemplateValues — The patcher should receive a plain data object        ║
 * ║ rather than mutate the row model.                                          ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ CLASS: OutputLoadRow                                                       │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • values: object keyed by template field key.
 *
 * Returns
 *   • An OutputLoadRow instance.
 *
 * Why / How
 *   Keeping rows behind a tiny model gives the XML patcher one predictable
 *   interface instead of depending on factory internals.
 *   It stores the logical field-value object and returns a shallow copy when
 *   writing begins.
 */
export class OutputLoadRow {
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ CONSTRUCTOR: constructor                                                   │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • values: completed template field values.
   *
   * Returns
   *   • A new OutputLoadRow.
   *
   * Why / How
   *   Each output row needs a single home for its completed logical values.
   *   It stores the values object without modifying it.
   */
  constructor(values) {
    this.values = values;
  }

  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ METHOD: toTemplateValues                                                   │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • None.
   *
   * Returns
   *   • A plain field-to-value object.
   *
   * Why / How
   *   The patcher should receive a plain data object rather than mutate the
   *   row model.
   *   It returns a shallow copy of the stored values.
   */
  toTemplateValues() {
    return { ...this.values };
  }
}
