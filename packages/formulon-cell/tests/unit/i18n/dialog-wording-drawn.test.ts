import { describe, expect, it } from 'vitest';
import { de } from '../../../src/i18n/strings/de.js';
import { en } from '../../../src/i18n/strings/en.js';
import { fr } from '../../../src/i18n/strings/fr.js';
import { ko } from '../../../src/i18n/strings/ko.js';
import { th } from '../../../src/i18n/strings/th.js';

describe('dialog wording', () => {
  it.each([
    ['en Calculation R1C1', en.iterativeDialog.r1c1, 'Use R1C1 reference style'],
    ['fr Find Within', fr.findReplace.within, 'Dans :'],
    [
      'ko Text to Columns',
      ko.mac.textToColumnsWizard.fixedPrompt,
      '각 필드의 너비(열 구분선)를 지정합니다.',
    ],
    ['th Go To Special', th.goToDialog.kindPrecedents, 'เซลล์ที่ถูกอ้าง'],
    ['en Page Setup', en.pageSetup.pageOrder, 'Page Order'],
    ['de Page Setup', de.pageSetup.blackAndWhite, 'Schwarzweißdruck'],
  ])('reads %s as the desktop draws it', (_, ours, drawn) => {
    expect(ours).toBe(drawn);
  });
});
