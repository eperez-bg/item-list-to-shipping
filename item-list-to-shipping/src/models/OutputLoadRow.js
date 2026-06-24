export class OutputLoadRow {
  constructor(values) {
    this.values = values;
  }

  toTemplateValues() {
    return { ...this.values };
  }
}
