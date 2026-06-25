export const ERROR_VALUE = "ERROR";

export class ImportIssue {
  constructor({ type, location, field, message, resolution }) {
    this.type = type;
    this.location = location;
    this.field = field;
    this.message = message;
    this.resolution = resolution;
  }
}

export class ImportIssueReporter {
  constructor(existingIssues = []) {
    this.issues = existingIssues;
  }

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
