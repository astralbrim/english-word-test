import { Utils } from './utils';

export class TestSheet {
  constructor(private readonly testSheet: GoogleAppsScript.Spreadsheet.Sheet) {
    this.utils = new module.exports.Utils();
  }
  utils: Utils;

  private setFormat() {
    this.utils.clearContentsFormat(this.testSheet);
    this.utils.removeImages(this.testSheet);
    this.testSheet
      .getRange('A1:C1')
      .merge()
      .setFontSize(30)
      .setWrap(true)
      .setVerticalAlignment('middle');
    this.testSheet
      .getRange('D1')
      .setHorizontalAlignment('right')
      .setFontSize(24)
      .setVerticalAlignment('bottom');
  }

  public toTest(
    title: string,
    questionNums: number,
    words: { english: string; japanese: string; number: string }[],
    webhookURL: string,
    mode: string,
  ) {
    this.setFormat();
    this.testSheet.getRange('A1').setValue(title.replaceAll(',', ''));
    this.testSheet.getRange('D1').setValue(`/${questionNums}`);
    this.testSheet
      .getRange('E1')
      .setValue('答え')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('bottom')
      .setFontSize(20);
    const questions = this.utils.getRandom(words, questionNums);
    let str = 'word=';
    questions.forEach((question) => {
      str = str.concat(question.number + ',');
    });
    str = str.slice(0, -1);
    const str0 = `${str}%26mode=E`;
    const str1 = `${str}%26mode=J`;
    this.testSheet.insertImage(
      this.utils.createQrCode(this.utils.toURL(webhookURL, 'test', str0)),
      4,
      1,
      410,
      0,
    );
    this.testSheet.insertImage(
      this.utils.createQrCode(this.utils.toURL(webhookURL, 'test', str1)),
      4,
      1,
      180,
      0,
    );
    this.testSheet.insertImage(
      this.utils.createQrCode(this.utils.toURL(webhookURL, 'answer', str)),
      5,
      1,
      30,
      0,
    );
    const images = this.testSheet.getImages();
    images.forEach((image) => image.setWidth(230).setHeight(230));
    const row = 3;
    this.testSheet.setRowHeights(row, questionNums / 2 + 1, 100);
    const testRange = this.testSheet.getRange(
      row,
      1,
      Math.round(questionNums / 2 + 1),
      4,
    );
    testRange.setFontSize(20);
    this.testSheet
      .getRange(row, 2, questionNums / 2 + 1)
      .setHorizontalAlignment('left');
    this.testSheet
      .getRange(row, 5, questionNums / 2 + 1)
      .setHorizontalAlignment('left');
    testRange.setBorder(true, true, true, true, true, true);
    let isEnglishToJapanese;
    if (mode) isEnglishToJapanese = mode[0] === 'E';
    else isEnglishToJapanese = true;

    testRange.getCell(1, 1).setValue('No');
    testRange.getCell(1, 3).setValue('No');
    for (let i = 2; i < questionNums + 2; i++) {
      if (i <= (questionNums + 1) / 2 + 1) {
        testRange.getCell(i, 1).setValue(questions[i - 2].number);
        testRange
          .getCell(i, 2)
          .setValue(
            isEnglishToJapanese
              ? questions[i - 2].english
              : questions[i - 2].japanese,
          )
          .setVerticalAlignment('top');
      } else {
        testRange
          .getCell(i - questionNums / 2, 3)
          .setValue(questions[i - 2].number);
        testRange
          .getCell(i - questionNums / 2, 4)
          .setValue(
            isEnglishToJapanese
              ? questions[i - 2].english
              : questions[i - 2].japanese,
          )
          .setVerticalAlignment('top');
      }
    }
  }
}
module.exports.TestSheet = TestSheet;
