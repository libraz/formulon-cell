import { describe, expect, it } from 'vitest';
import { ko } from '../../../src/i18n/strings/ko.js';
import { zh } from '../../../src/i18n/strings/zh.js';
import { macRibbonLabelForCommand } from '../../../src/toolbar/ribbon/mac/labels.js';
import { createTableMenuFactories } from '../../../src/toolbar/ribbon/menus/table.js';
import type { ToolbarLang } from '../../../src/toolbar/ribbon-model.js';

/** The ribbon draws these zero-width spaces between the words of some labels. */
const ZWSP = String.fromCodePoint(0x200b);
const NBSP = String.fromCodePoint(0xa0);

/** Text of one item of the Table Design tab's Refresh menu. */
const tableRefreshItem = (lang: ToolbarLang, value: string): string | null | undefined =>
  createTableMenuFactories(lang, 'Refresh')
    .createTableRefreshMenu('tableRefresh')
    .querySelector(`[data-table-refresh="${value}"] .fc-tb__menu-item__text`)?.textContent;

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
    [
      'en Apply Names',
      macRibbonLabelForCommand('mac.formulas.defineName.apply', 'en'),
      'Apply Names...',
    ],
    [
      'ja Apply Names',
      macRibbonLabelForCommand('mac.formulas.defineName.apply', 'ja'),
      '名前の引用...',
    ],
    [
      'de Apply Names',
      macRibbonLabelForCommand('mac.formulas.defineName.apply', 'de'),
      'Namen übernehmen...',
    ],
    [
      'ja Horizontal Text Box',
      macRibbonLabelForCommand('mac.insert.textBox.horizontal', 'ja'),
      '横書きテキスト ボックスの描画',
    ],
    [
      'ko Horizontal Text Box',
      macRibbonLabelForCommand('mac.insert.textBox.horizontal', 'ko'),
      '가로 텍스트 상자 그리기',
    ],
    [
      'zh Horizontal Text Box',
      macRibbonLabelForCommand('mac.insert.textBox.horizontal', 'zh'),
      '绘制横排文本框',
    ],
    [
      'ja Vertical Text Box',
      macRibbonLabelForCommand('mac.insert.textBox.vertical', 'ja'),
      '縦書きテキスト ボックス',
    ],
    [
      'ja Next Comment',
      macRibbonLabelForCommand('mac.review.nextComment', 'ja'),
      `次の${ZWSP}${ZWSP}コメント`,
    ],
    [
      'ja Previous Comment',
      macRibbonLabelForCommand('mac.review.previousComment', 'ja'),
      `前の${ZWSP}${ZWSP}コメント`,
    ],
    ['en Analysis Tools', macRibbonLabelForCommand('mac.data.analysis', 'en'), 'Analysis Tools'],
    ['de Analysis Tools', macRibbonLabelForCommand('mac.data.analysis', 'de'), 'Analysetools'],
    ['zh Analysis Tools', macRibbonLabelForCommand('mac.data.analysis', 'zh'), '分析工具'],
    [
      'ja Text to Columns',
      macRibbonLabelForCommand('mac.data.textToColumns', 'ja'),
      `区切り位置${ZWSP}${ZWSP}`,
    ],
    [
      'zh Text to Columns',
      macRibbonLabelForCommand('mac.data.textToColumns', 'zh'),
      `分列${ZWSP}${ZWSP}`,
    ],
    ['de View Macros', macRibbonLabelForCommand('mac.view.macros', 'de'), 'Makros'],
    ['fr View Macros', macRibbonLabelForCommand('mac.view.macros', 'fr'), 'Macros'],
    ['th View Macros', macRibbonLabelForCommand('mac.view.macros', 'th'), 'แสดงแมโคร'],
    [
      'de Page Setup',
      macRibbonLabelForCommand('mac.page.pageSetup', 'de'),
      `Seiteneinrichtung${NBSP}…`,
    ],
    ['de Refresh Status', tableRefreshItem('de', 'refresh-status'), 'Status aktualisieren'],
    ['fr Refresh Status', tableRefreshItem('fr', 'refresh-status'), 'Actualiser l’état'],
    [
      'en Connection Properties',
      tableRefreshItem('en', 'connection-properties'),
      'Connection Properties...',
    ],
    [
      'ja Connection Properties',
      tableRefreshItem('ja', 'connection-properties'),
      '接続のプロパティ...',
    ],
  ])('reads %s as the desktop draws it', (_, ours, drawn) => {
    expect(ours).toBe(drawn);
  });
});
