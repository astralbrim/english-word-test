export class Utils {
  public createQrCode(code_data) {
    const url =
      'https://chart.googleapis.com/chart?chs=100x100&cht=qr&chl=' + code_data;
    const ajax = UrlFetchApp.fetch(url, { method: 'get' });
    console.log(ajax.getBlob());
    return ajax.getBlob();
  }
  public getRandom<T>(arr: T[], n: number): T[] {
    let result = new Array(n),
      len = arr.length,
      taken = new Array(len);
    if (n > len)
      throw new RangeError('getRandom: more elements taken than available');
    while (n--) {
      const x = Math.floor(Math.random() * len);
      result[n] = arr[x in taken ? taken[x] : x];
      taken[x] = --len in taken ? taken[len] : len;
    }
    return result;
  }
  public getPdf(spreadSheet: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    const sheet = spreadSheet.getSheetByName('単語テスト');
    const newSheet = SpreadsheetApp.create('印刷用');
    sheet.copyTo(newSheet);
    newSheet.deleteSheet(newSheet.getSheets()[0]);
    const pdf = newSheet.getAs('application/pdf');
    pdf.setName('英単語テスト.pdf');
    const folderRoot = DriveApp.getFileById(spreadSheet.getId())
      .getParents()
      .next();
    const tests = folderRoot.getFilesByName('英単語テスト.pdf');
    while (tests.hasNext()) folderRoot.removeFile(tests.next());
    folderRoot.createFile(pdf);
    return folderRoot.getFiles().next().getUrl();
  }
  public clearContentsFormat(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    sheet.clearContents().clearFormats();
  }
  public removeImages(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    sheet.getImages().forEach((image) => image.remove());
  }
  public getCheckedUnitList(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
  ): string[] {
    const units = sheet.getRange(2, 4, 100);
    const checkBoxes = sheet.getRange(2, 5, 100);

    const testUnits = Array<string>();
    for (let row = 1; row < 100; row++) {
      const checkBox = checkBoxes.getCell(row, 1).getValue();
      const unit = units.getCell(row, 1).getValue();
      if (checkBox) testUnits.push(unit);
    }
    return testUnits;
  }
  public getTest(
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
    title: string,
    questionNums: number,
    words: { english: string; japanese: string; number: string }[],
    mode?: string,
  ) {
    const testSheet = spreadsheet.getSheetByName('単語テスト');
    const controlSheet = spreadsheet.getSheetByName('操作');
    testSheet
      .getRange('A1:C1')
      .merge()
      .setValue(title.replaceAll(',', ''))
      .setFontSize(30)
      .setWrap(true)
      .setVerticalAlignment('middle');
    testSheet
      .getRange('D1')
      .setValue(`/${questionNums}`)
      .setHorizontalAlignment('right')
      .setFontSize(24)
      .setVerticalAlignment('bottom');
    const webhookURL = controlSheet.getRange('G2').getValue();
    const questions = this.getRandom(words, questionNums);
    let str = '';
    questions.forEach((question) => {
      str = str.concat(question.number + ',');
    });
    str = str.slice(0, -1);
    const data = `${webhookURL}?word=${str}`;
    controlSheet.getRange('A1').setValue(data);
    testSheet.insertImage(this.createQrCode(data), 4, 1, 410, 0);
    const image = testSheet.getImages()[0];
    image.setHeight(230);
    image.setWidth(230);
    const row = 3;
    testSheet.setRowHeights(row, questionNums / 2 + 1, 100);
    const testRange = testSheet.getRange(
      row,
      1,
      Math.round(questionNums / 2 + 1),
      4,
    );
    testRange.setFontSize(20);
    testSheet
      .getRange(row, 2, questionNums / 2 + 1)
      .setHorizontalAlignment('left');
    testSheet
      .getRange(row, 5, questionNums / 2 + 1)
      .setHorizontalAlignment('left');
    testRange.setBorder(true, true, true, true, true, true);
    const isEnglishToJapanese = mode ? mode === '英' : true;
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

module.exports.Utils = Utils;
