import { Utils } from './utils';

export class ListSheet {
  constructor(private readonly listSheet: GoogleAppsScript.Spreadsheet.Sheet) {
    this.utils = new module.exports.Utils();
  }
  utils: Utils;

  public getWords(
    testUnits: string[],
    grade: number,
  ): Array<{ english: string; japanese: string; number: string }> {
    const unitNameRow = this.listSheet.getRange(
      3,
      this.utils.getColumn(grade, 2),
      2000,
    );
    const words = Array<{
      english: string;
      japanese: string;
      number: string;
    }>();
    for (let i = 0; i < testUnits.length; i++) {
      const finder = unitNameRow.createTextFinder(testUnits[i]);
      const cell = finder.findNext().getCell(1, 1);
      const row = cell.getRow();
      const column = cell.getColumn();
      const range = this.listSheet.getRange(row, column - 1, 100, 4);
      for (let j = 1; ; j++) {
        const english = range.getCell(j, 3).getValue();
        const japanese = range.getCell(j, 4).getValue();
        const number = range.getCell(j, 1).getValue();
        if (english === '' && japanese === '') break;
        words.push({ english, japanese, number });
      }
    }
    return words;
  }
}
module.exports.ListSheet = ListSheet;