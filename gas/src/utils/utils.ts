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

  public generateTitle(grade: number, testUnits: string[]): string {
    return `英単語テスト     ${grade}年生      ${testUnits.map(
      (el) => el + '  ',
    )}`;
  }

  public clearContentsFormat(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    sheet.clearContents().clearFormats();
  }

  public removeImages(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
    sheet.getImages().forEach((image) => image.remove());
  }
  public getColumn(grade: number, column: number): number {
    return column + 4 * (grade - 1);
  }
}

module.exports.Utils = Utils;
