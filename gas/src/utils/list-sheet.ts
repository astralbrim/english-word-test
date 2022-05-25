import { Utils } from './utils';

export class ListSheet {
  constructor(private readonly listSheet: GoogleAppsScript.Spreadsheet.Sheet) {
    this.utils = new module.exports.Utils();
  }
  utils: Utils;

  public getWordsByUnit(
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

  public getWordsById(
    wordIds: string[],
  ): Array<{ english: string; japanese: string; number: string }> {
    const words = Array<{
      english: string;
      japanese: string;
      number: string;
    }>();
    wordIds.forEach((number) => {
      const finder1 = this.listSheet
        .getRange(3, 1, 2000)
        .createTextFinder(number);
      const finder2 = this.listSheet
        .getRange(3, 5, 2000)
        .createTextFinder(number);
      const finder3 = this.listSheet
        .getRange(3, 9, 2000)
        .createTextFinder(number);
      let col, row;
      const next1 = finder1.findNext();
      const next2 = finder2.findNext();
      const next3 = finder3.findNext();

      switch (true) {
        case !!next1:
          col = next1.getColumn();
          row = next1.getRow();
          break;
        case !!next2:
          col = next2.getColumn();
          row = next2.getRow();
          break;
        case !!next3:
          col = next3.getColumn();
          row = next3.getRow();
          break;
      }

      const range = this.listSheet.getRange(row, col, 1, 4);
      const japanese = range.getCell(1, 4).getValue();
      const english = range.getCell(1, 3).getValue();
      words.push({ english, japanese, number });
    });
    return words;
  }
}
module.exports.ListSheet = ListSheet;
