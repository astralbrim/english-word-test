import { UnitNamesEntity } from '../../server/src/scrape/unit-name.entity';
import {
  UnitsEntity,
} from '../../server/src/scrape/word.entity';


// noinspection JSUnusedLocalSymbols
function updateUnitList() {
  const baseUrl =
    'https://a587-240d-1a-81f-c100-1063-636f-e968-4298.ngrok.io/api';
  const spreadsheet = SpreadsheetApp.getActive();
  const controlSheet = spreadsheet.getSheetByName('操作');

  const gradeCell = controlSheet.getRange('C2');
  const grade = gradeCell.getValue()[0];
  const res = UrlFetchApp.fetch(`${baseUrl}/units/${grade}`);
  const units: UnitNamesEntity = JSON.parse(res.getContentText());
  const unitNameRange = controlSheet.getRange(2, 4, 100);
  const checkBoxRange = controlSheet.getRange(2, 5, 100);
  console.log(units.unitNames);
  units.unitNames.forEach((unit, index) => {
    unitNameRange.getCell(index + 1, 1).setValue(unit);
    checkBoxRange.getCell(index + 1, 1).insertCheckboxes();
  });
}
// noinspection JSUnusedLocalSymbols
function updateWordList() {
  const baseUrl =
    'https://a587-240d-1a-81f-c100-1063-636f-e968-4298.ngrok.io/api';
  const MAX = 10000;
  const spreadsheet = SpreadsheetApp.getActive();
  const listSheet = spreadsheet.getSheetByName('単語一覧');
  const controlSheet = spreadsheet.getSheetByName('操作');

  const gradeCell = controlSheet.getRange('C2');
  const grade = gradeCell.getValue()[0];

  const unitRange = listSheet.getRange(3, 1 + (grade - 1) * 3, MAX);
  const englishRange = listSheet.getRange(3, 2 + (grade - 1) * 3, MAX);
  const japaneseRange = listSheet.getRange(3, 3 + (grade - 1) * 3, MAX);

  const res = UrlFetchApp.fetch(`${baseUrl}/words/${grade}`);
  const units: UnitsEntity = JSON.parse(res.getContentText());
  let count = 0;
  units.units.forEach((unit, index) => {
    const unitName = unit.unitName;
    unitRange.getCell(index + count + 1, 1).setValue(unitName);
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
