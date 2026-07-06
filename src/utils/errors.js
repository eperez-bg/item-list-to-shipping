/*
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║ FILE: errors.js                                                            ║
 * ╠══════════════════════════════════════════════════════════════════════════╣
 * ║ RESPONSIBILITY                                                             ║
 * ║ Defines errors meant to be shown to the user and a small assertion         ║
 * ║ helper that throws them consistently.                                      ║
 * ║                                                                            ║
 * ║ WHY THIS FILE EXISTS                                                       ║
 * ║ Expected user-correctable problems need a distinct error type so the UI    ║
 * ║ can display a clean message instead of treating them as programmer         ║
 * ║ crashes.                                                                   ║
 * ║                                                                            ║
 * ║ MAIN CLASSES / FUNCTIONS                                                   ║
 * ║ • UserFacingError — Expected user-correctable problems need a distinct     ║
 * ║ error type so the UI can display a clean message instead of treating       ║
 * ║ them as programmer crashes.                                                ║
 * ║ • assertUser — Validation code should have a concise way to stop with a    ║
 * ║ user-friendly error when a required condition is false.                    ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ CLASS: UserFacingError                                                     │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • message: human-readable explanation.
 *
 * Returns
 *   • Error instance intended for UI display.
 *
 * Why / How
 *   Expected user-correctable problems need a distinct error type so the UI
 *   can display a clean message instead of treating them as programmer
 *   crashes.
 *   It subclasses Error and assigns a stable name.
 */
export class UserFacingError extends Error {
  /**
   * ┌──────────────────────────────────────────────────────────────────────────┐
   * │ CONSTRUCTOR: constructor                                                   │
   * └──────────────────────────────────────────────────────────────────────────┘
   *
   * Parameters
   *   • message: user-facing text.
   *
   * Returns
   *   • A configured UserFacingError instance.
   *
   * Why / How
   *   The specialized error must behave like a normal JavaScript Error while
   *   preserving a recognizable name.
   *   It calls the parent constructor and changes the name property.
   */
  constructor(message) {
    super(message);
    this.name = "UserFacingError";
  }
}

/**
 * ┌──────────────────────────────────────────────────────────────────────────┐
 * │ FUNCTION: assertUser                                                       │
 * └──────────────────────────────────────────────────────────────────────────┘
 *
 * Parameters
 *   • condition and displayed message.
 *
 * Returns
 *   • No returned value when valid; otherwise throws UserFacingError.
 *
 * Why / How
 *   Validation code should have a concise way to stop with a user-friendly
 *   error when a required condition is false.
 *   It throws UserFacingError only when condition is falsy.
 */
export function assertUser(condition, message) {
  if (!condition) {
    throw new UserFacingError(message);
  }
}
