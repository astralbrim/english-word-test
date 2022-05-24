import { UnitsEntity } from '../../server/src/scrape/word.entity';
import DoGet = GoogleAppsScript.Events.DoGet;

// noinspection JSUnusedLocalSymbols
function init(option?: { setting?: boolean; unitList?: boolean }) {
  const spreadSheet = SpreadsheetApp.getActive();
  const controlSheet = spreadSheet.getSheetByName('操作');
  if (!option || option.setting) {
    controlSheet.getRange('C2').setValue('');
    controlSheet.getRange('C4').setValue('');
    controlSheet.getRange('C6').setValue('');
  }
  if (!option || option.unitList) {
    controlSheet.getRange(2, 4, 100).clear();
    controlSheet.getRange(2, 5, 100).clearDataValidations().clear();
  }
}
// noinspection JSUnusedLocalSymbols
function updateUnitList() {
  const spreadsheet = SpreadsheetApp.getActive();
  const controlSheet = spreadsheet.getSheetByName('操作');
  const listSheet = spreadsheet.getSheetByName('単語一覧');
  const grade = controlSheet.getRange('C2').getValue()[0];
  init({ unitList: true, setting: false });

  const unitList = listSheet.getRange(3, 2 + 4 * (grade - 1), 2000);
  const unitNameRange = controlSheet.getRange(2, 4, 100);
  const checkBoxRange = controlSheet.getRange(2, 5, 100);

  const array = Array<string>();
  for (let col = 1; col <= 2000; col++) {
    const cell = unitList.getCell(col, 1);
    if (!cell.isBlank()) {
      array.push(cell.getValue());
    }
  }

  array.forEach((el, index) => {
    unitNameRange.getCell(index + 1, 1).setValue(el);
    checkBoxRange.getCell(index + 1, 1).insertCheckboxes();
  });
}
// noinspection JSUnusedLocalSymbols
function updateWordList() {
  const baseUrl =
    'https://42d1-240d-1a-81f-c100-f5cb-aad2-f726-b157.ngrok.io/api';
  const MAX = 10000;
  const spreadsheet = SpreadsheetApp.getActive();
  const listSheet = spreadsheet.getSheetByName('単語一覧');
  const controlSheet = spreadsheet.getSheetByName('操作');

  const gradeCell = controlSheet.getRange('C2');
  const grade = gradeCell.getValue()[0];

  const unitRange = listSheet.getRange(3, 2 + (grade - 1) * 4, MAX);
  const englishRange = listSheet.getRange(3, 3 + (grade - 1) * 4, MAX);
  const japaneseRange = listSheet.getRange(3, 4 + (grade - 1) * 4, MAX);

  const res = UrlFetchApp.fetch(`${baseUrl}/words/${grade}`);
  const units: UnitsEntity = JSON.parse(res.getContentText());
  let count = 0;
  units.units.forEach((unit, index) => {
    unitRange.getCell(index + count + 1, 1).setValue(unit.unitName);
    unit.words.forEach((word) => {
      const index2 = index + count;
      const english = word.english;
      const japanese = word.japanese;
      englishRange.getCell(index2 + 1, 1).setValue(english);
      japaneseRange.getCell(index2 + 1, 1).setValue(japanese);
      count = count + 1;
    });
  });
}
// noinspection JSUnusedLocalSymbols
function generateTestFromQrCode(wordIds: string[]) {
  const Utils = module.exports.Utils;
  const utils = new Utils();
  const spreadsheet = SpreadsheetApp.getActive();
  const listSheet = spreadsheet.getSheetByName('単語一覧');
  const controlSheet = spreadsheet.getSheetByName('操作');
  const testSheet = spreadsheet.getSheetByName('単語テスト');
  utils.clearContentsFormat(testSheet);
  utils.removeImages(testSheet);
  wordIds.forEach(() => console.log());
  const words = Array<{ english: string; japanese: string; number: string }>();
  wordIds.forEach((number) => {
    const finder1 = listSheet.getRange(3, 1, 2000).createTextFinder(number);
    const finder2 = listSheet.getRange(3, 5, 2000).createTextFinder(number);
    const finder3 = listSheet.getRange(3, 9, 2000).createTextFinder(number);
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

    const range = listSheet.getRange(row, col, 1, 4);
    const japanese = range.getCell(1, 4).getValue();
    const english = range.getCell(1, 3).getValue();
    words.push({ english, japanese, number });
  });
  const questionNums = wordIds.length;
  const title = '英単語テスト';
  utils.getTest(spreadsheet, title, questionNums, words);
  controlSheet.getRange('F2').setValue(utils.getPdf(spreadsheet));
}
// noinspection JSUnusedLocalSymbols
function generateTest() {
  const Utils = module.exports.Utils;
  const utils = new Utils();
  const spreadsheet = SpreadsheetApp.getActive();
  const listSheet = spreadsheet.getSheetByName('単語一覧');
  const controlSheet = spreadsheet.getSheetByName('操作');
  const testSheet = spreadsheet.getSheetByName('単語テスト');

  utils.clearContentsFormat(testSheet);
  utils.removeImages(testSheet);
  const gradeCell = controlSheet.getRange('C2');
  const modeCell = controlSheet.getRange('C4');
  const questionNumsCell = controlSheet.getRange('C6');
  const grade = gradeCell.getValue()[0];
  const mode = modeCell.getValue() as string;
  // バリデーション
  if (!grade || !mode) {
    if (!grade) gradeCell.setBackground('pink');
    if (!mode) modeCell.setBackground('pink');
    Browser.msgBox('必要な情報が入力されていません');
    return;
  } else {
    gradeCell.setBackground('white');
    modeCell.setBackground('white');
  }
  // テストのユニットリスト
  const testUnits = utils.getCheckedUnitList(controlSheet);

  const unitNameRow = listSheet.getRange(3, 2 + 4 * (grade - 1), 2000);
  const words = Array<{ english: string; japanese: string; number: string }>();
  for (let i = 0; i < testUnits.length; i++) {
    const finder = unitNameRow.createTextFinder(testUnits[i]);
    const cell = finder.findNext().getCell(1, 1);
    const row = cell.getRow();
    const column = cell.getColumn();
    const range = listSheet.getRange(row, column - 1, 100, 4);
    for (let j = 1; ; j++) {
      const english = range.getCell(j, 3).getValue();
      const japanese = range.getCell(j, 4).getValue();
      const number = range.getCell(j, 1).getValue();
      if (english === '' && japanese === '') break;
      words.push({ english, japanese, number });
    }
  }
  const questionNums = (() => {
    if (questionNumsCell.isBlank()) return words.length;
    else return questionNumsCell.getValue();
  })();
  const title = `英単語テスト     ${grade}年生      ${testUnits.map(
    (el) => el + '  ',
  )}`;
  utils.getTest(spreadsheet, title, questionNums, words, mode);
  controlSheet.getRange('F2').setValue(utils.getPdf(spreadsheet));
}
// noinspection JSUnusedLocalSymbols
function doGet(e: DoGet) {
  const wordIds = e.parameters['word'].toString().split(',');
  generateTestFromQrCode(wordIds);
  return ContentService.createTextOutput(
    '更新が完了しました。スプレッドシートの操作シートのpdfを印刷してください。',
  );
}
