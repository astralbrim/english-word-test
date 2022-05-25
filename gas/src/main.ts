import { UnitsEntity } from '../../server/src/scrape/word.entity';
import DoGet = GoogleAppsScript.Events.DoGet;
import { TestSheet } from './utils/test-sheet';
import { ControlSheet } from './utils/control-sheet';
import { ListSheet } from './utils/list-sheet';
import { Utils } from './utils/utils';
import {RequestType} from "./utils/request-type";
function getSpreadSheet(): {
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  listSheet: ListSheet;
  controlSheet: ControlSheet;
  testSheet: TestSheet;
} {
  const spreadsheet = SpreadsheetApp.getActive();
  const listSheet: ListSheet = new module.exports.ListSheet(
    spreadsheet.getSheetByName('単語一覧'),
  );
  const controlSheet: ControlSheet = new module.exports.ControlSheet(
    spreadsheet.getSheetByName('操作'),
  );
  const testSheet: TestSheet = new module.exports.TestSheet(
    spreadsheet.getSheetByName('単語テスト'),
  );
  return { spreadsheet, listSheet, controlSheet, testSheet };
}
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
function generateTestFromQrCode(wordIds: string[], mode: string) {
  const Utils = module.exports.Utils;
  const utils: Utils = new Utils();
  const {
    spreadsheet,
    listSheet,
    controlSheet,
    testSheet,
  }: {
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
    listSheet: ListSheet;
    controlSheet: ControlSheet;
    testSheet: TestSheet;
  } = getSpreadSheet();
  const { webhookURL } = controlSheet.getInfo();
  wordIds.forEach(() => console.log());
  testSheet.toTest(
    '英単語テスト',
    wordIds.length,
    listSheet.getWordsById(wordIds),
    webhookURL,
    mode,
  );
  controlSheet.setPdfUrl(utils.getPdf(spreadsheet));
}
// noinspection JSUnusedLocalSymbols
function generateTest() {
  const { spreadsheet, listSheet, controlSheet, testSheet } = getSpreadSheet();
  const utils: Utils = new module.exports.Utils();

  const { gradeCell, modeCell, grade, mode, webhookURL } =
    controlSheet.getInfo();
  // バリデーション
  const { grade: isGradeBlank, mode: isModeBlank } =
    controlSheet.checkValidation(gradeCell, modeCell);
  if (isGradeBlank) controlSheet.changeCellColor(gradeCell, 'pink');
  if (isModeBlank) controlSheet.changeCellColor(modeCell, 'pink');
  if (isGradeBlank || isModeBlank) {
    controlSheet.showDialog('必要な情報が入力されていません');
    return;
  } else {
    controlSheet.changeCellColor(gradeCell, 'white');
    controlSheet.changeCellColor(modeCell, 'white');
  }

  // テストにする単元名
  const testUnits = controlSheet.getCheckedUnitList();
  // テストにする単語
  const words = listSheet.getWordsByUnit(testUnits, grade);
  // 問題数
  const questionNums = controlSheet.getQuestionNums(words);
  // タイトルの作成
  const title = utils.generateTitle(grade, testUnits);
  //テストの作成
  testSheet.toTest(title, questionNums, words, webhookURL, mode);
  // pdfに変換
  const url = utils.getPdf(spreadsheet);
  controlSheet.setPdfUrl(url);
}
// noinspection JSUnusedLocalSymbols
function getAnswer(wordIds: string[]): {
  number: string, english: string, japanese: string
}[] {
  const {listSheet} = getSpreadSheet();
  return  listSheet.getWordsById(wordIds);
}
// noinspection JSUnusedLocalSymbols
function doGet(e: DoGet) {
  if(!e.parameters['requestType']) return;
  const requestType = e.parameters['requestType'].toString() as RequestType;
  let wordIds;
  switch (requestType) {
    case "test":
      wordIds = e.parameters['word'].toString().split(',');
      const mode = e.parameters['mode'].toString();
      generateTestFromQrCode(wordIds, mode);
      return ContentService.createTextOutput(
          '更新が完了しました。スプレッドシートの操作シートのpdfを印刷してください。',
      );
    case "answer":
      wordIds = e.parameters['word'].toString().split(',');
      const answer = getAnswer(wordIds);
      const template =  HtmlService.createHtmlOutputFromFile("templates");
      template.append("<ul>")
      answer.forEach(
          (elem) => {
            template.append(`<li>${elem.number}  ${elem.japanese}  ${elem.english}</li>`)
          }
      )
      template.append("</ul>")
      return template;
  }

}
