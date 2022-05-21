export class UnitNamesEntity {
  constructor() {
    this.unitNames = [];
  }
  public unitNames: string[];

  addUnitName(unitName: string) {
    this.unitNames.push(unitName);
  }
}
