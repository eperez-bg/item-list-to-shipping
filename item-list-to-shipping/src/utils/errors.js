export class UserFacingError extends Error {
  constructor(message) {
    super(message);
    this.name = "UserFacingError";
  }
}

export function assertUser(condition, message) {
  if (!condition) {
    throw new UserFacingError(message);
  }
}
