import { describe, expect, it } from 'vitest';
import {
  macPivotOptionsDe,
  macPivotOptionsEn,
  macPivotOptionsFr,
  macPivotOptionsJa,
  macPivotOptionsKo,
  macPivotOptionsTh,
  macPivotOptionsZh,
} from '../../../src/i18n/strings/mac/pivot-options.js';

describe('PivotTable Options wording', () => {
  it.each([
    ['en', macPivotOptionsEn.display.printTitles, 'Page, row and column headings'],
    ['en', macPivotOptionsEn.data.backgroundQuery, 'Run queries in the background'],
    ['ja', macPivotOptionsJa.display.printTitles, '印刷タイトルを設定する'],
    [
      'de',
      macPivotOptionsDe.display.itemLabels,
      'Elementbeschriftungen, wenn sich keine Felder in der Tabelle befinden',
    ],
    [
      'fr',
      macPivotOptionsFr.display.itemLabels,
      "Étiquettes d'élément si le tableau ne contient pas de champs",
    ],
    ['ko', macPivotOptionsKo.data.savePassword, '파일과 함께 암호 저장'],
    ['zh', macPivotOptionsZh.data.savePassword, '保存密码及文件'],
    ['th', macPivotOptionsTh.display.noDataRows, 'บนแถว'],
  ])('labels the %s dialog as the desktop draws it', (_, ours, drawn) => {
    expect(ours).toBe(drawn);
  });
});
