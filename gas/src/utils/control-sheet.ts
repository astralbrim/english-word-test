import { Utils } from './utils';

export class ControlSheet {
  constructor(
    private readonly controlSheet: GoogleAppsScript.Spreadsheet.Sheet,
  ) {
    this.utils = new module.exports.Utils();
  }
  private utils: Utils;

  public getInfo(): {
    gradeCell: GoogleAppsScript.Spreadsheet.Range;
    modeCell: GoogleAppsScript.Spreadsheet.Range;
    grade: number;
    mode: string;
    questionNums: number;
    webhookURL: string;
  } {
    const gradeCell = this.controlSheet.getRange('C2');
    const modeCell = this.controlSheet.getRange('C4');
    const questionNumsCell = this.controlSheet.getRange('C6');
    const webhookURLCell = this.controlSheet.getRange('G2');
    return {
      gradeCell,
      modeCell,
      grade: Number(gradeCell.getValue()[0]),
      mode: modeCell.getValue(),
      questionNums: Number(questionNumsCell.getValue()),
      webhookURL: webhookURLCell.getValue(),
    };
  }
  public checkValidation(
    gradeCell: GoogleAppsScript.Spreadsheet.Range,
    modeCell: GoogleAppsScript.Spreadsheet.Range,
  ): { grade: boolean; mode: boolean } {
    const grade = gradeCell.getValue();
    const mode = modeCell.getValue();
    return {
      grade: !grade,
      mode: !mode,
    };
  }
  public changeCellColor(
    range: GoogleAppsScript.Spreadsheet.Range,
    color: 'pink' | 'white',
  ) {
    range.setBackground(color);
  }
  public showDialog(str: string) {
    Browser.msgBox(str);
  }

  public getCheckedUnitList(): string[] {
    const units = this.controlSheet.getRange(2, 4, 100);
    const checkBoxes = this.controlSheet.getRange(2, 5, 100);

    const testUnits = Array<string>();
    for (let row = 1; row < 100; row++) {
      const checkBox = checkBoxes.getCell(row, 1).getValue();
      const unit = units.getCell(row, 1).getValue();
      if (checkBox) testUnits.push(unit);
    }
    return testUnits;
  }
  public getQuestionNums(words: Array<any>): number {
    const questionNumsCell = this.controlSheet.getRange('C6');
    if (questionNumsCell.isBlank()) return words.length;
    else return Number(questionNumsCell.getValue());
  }

  public setPdfUrl(url: string): void {
    this.controlSheet.getRange('F2').setValue(url);
  }
}
module.exports.ControlSheet = ControlSheet;
