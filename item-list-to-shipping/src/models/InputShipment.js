export class InputShipment {
  constructor({ skids, issues = [] }) {
    this.skids = skids;
    this.issues = issues;
  }

  get itemCount() {
    return this.skids.reduce((total, skid) => total + skid.itemCount, 0);
  }
}
