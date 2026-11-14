import { describe, expect, it } from 'vitest';
import { ko } from '../../../src/i18n/strings/ko.js';
import { zh } from '../../../src/i18n/strings/zh.js';
import { macRibbonLabelForCommand } from '../../../src/toolbar/ribbon/mac/labels.js';

describe('menu wording', () => {
  it.each([
    ['zh Orientation clockwise', zh.ribbonMenu.orientationAngleClockwise, '顺时针斜排'],
    [
      'zh Orientation counterclockwise',
      zh.ribbonMenu.orientationAngleCounterclockwise,
      '逆时针斜排',
    ],
    ['ko File Properties', ko.backstage.properties, '속성...'],
    [
      'ja Remove Arrows',
      macRibbonLabelForCommand('mac.formulas.removeArrows.all', 'ja'),
      'トレース矢印の削除',
    ],
    [
      'ko Remove Arrows',
      macRibbonLabelForCommand('mac.formulas.removeArrows.all', 'ko'),
      '연결선 제거',
    ],
  ])('reads %s as the desktop draws it', (_, ours, drawn) => {
    expect(ours).toBe(drawn);
  });
});
