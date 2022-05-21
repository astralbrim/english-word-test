import { Injectable } from '@nestjs/common';
import puppeteer from 'puppeteer';
import { UnitEntity, UnitsEntity, WordEntity } from './word.entity';
import { url } from './url';
import { UnitNamesEntity } from './unit-name.entity';
@Injectable()
export class ScrapeService {
  async getWords(grade: string) {
    // ブラウザを開く
    const browser = await puppeteer.launch({ headless: true });
    // 新しいページを開く
    const page = await browser.newPage();
    // ページ遷移
    await page.setUserAgent(
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36',
    );
    await page.goto(url[grade]);
    // <input></input>の取得
    await page.waitForSelector('#checkbox0');
    const inputs = await page.$$('#checkbox0');
    let i = -1;
    await page.waitForSelector('button#word_btn');
    const units = new UnitsEntity();
    for (let input of inputs) {
      i = i + 1;
      let unit;
      try {
        unit = await this.scrapeWords(grade, page, i);
      } catch (e) {
        console.error('not found');
        i = i - 1;
        continue;
      }
      units.addUnit(unit);
      if (i > 2) break;
    }
    await browser.close();
    return units;
  }
  async scrapeWords(
    grade: string,
    page: puppeteer.Page,
    time: number,
  ): Promise<UnitEntity> {
    await page.waitForSelector('#checkbox0');
    let input = (await page.$$('#checkbox0'))[time];
    console.log(time);
    await input.click();
    // ページ遷移
    const button = (await page.$$('button#word_btn'))[0];
    await button.click();
    // pageの更新
    await page.reload();

    // Unit名の取得
    await page.waitForSelector(
      '#content > div.row.words > div.columns.view_name_frame > div',
    );
    const element = await page.$(
      '#content > div.row.words > div.columns.view_name_frame > div',
    );
    const unitName = await page.evaluate((el) => el.textContent, element);
    const unit = new UnitEntity(unitName);

    // 単語の取得
    await page.waitForSelector(
      '#content > div.row.words > div.columns.word_col',
    );
    const wordCols = await page.$$('#word_table > tbody');
    for (const wordCol of wordCols) {
      const englishElement = await wordCol.$('#words > td.word_name');
      const english = await page.evaluate(
        (el) => el.textContent,
        englishElement,
      );
      const japaneseElement = await wordCol.$('tr > td.means');
      const japanese = await page.evaluate(
        (el) => el.textContent,
        japaneseElement,
      );
      const word = new WordEntity(english, japanese);
      unit.addWord(word);
    }
    await page.goto(url[grade]);
    await page.reload();
    return unit;
  }
  async getUnitNames(grade: string): Promise<UnitNamesEntity> {
    // ブラウザを開く
    const browser = await puppeteer.launch({ headless: true });
    // 新しいページを開く
    const page = await browser.newPage();
    // ページ遷移
    await page.goto(url[grade]);
    await page.waitForSelector('#td_2');
    const elements = await page.$$('#td_2');
    const unitNames = new UnitNamesEntity();
    for (const element of elements) {
      const unitName = await page.evaluate((el) => el.textContent, element);
      unitNames.addUnitName(unitName);
    }
    return unitNames;
  }
}
