import { expect, type Page } from '@playwright/test';

import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type Platform = 'default' | 'mac';
type Locale = 'en' | 'ja' | 'de';
type Tab = 'tableDesign' | 'pivotAnalyze' | 'pivotDesign' | 'sparklineDesign';

const UNTITLED = (n: number): string[] => Array.from({ length: n }, () => '');

/** Group titles per contextual tab, in ribbon order. The Mac profile reads the desktop's Mac
 *  labels; the default profile keeps its own titles and leaves the PivotTable groups untitled. */
const TITLES: Record<Platform, Partial<Record<Locale, Record<Tab, string[]>>>> = {
  mac: {
    en: {
      tableDesign: [
        'Properties',
        'Tools',
        'External Table Data',
        'Table Style Options',
        'Table Styles',
      ],
      pivotAnalyze: [
        'PivotTable',
        'Active Field',
        'Group',
        'Filter',
        'Data',
        'Actions',
        'Calculations',
        'Tools',
        'Show',
      ],
      pivotDesign: ['Layout', 'PivotTable Style Options', 'PivotTable Styles'],
      sparklineDesign: ['Sparkline', 'Type', 'Show', 'Style', 'Group'],
    },
    ja: {
      tableDesign: [
        'プロパティ',
        'ツール',
        '外部テーブル データ',
        'テーブル スタイル オプション',
        '表のスタイル',
      ],
      pivotAnalyze: [
        'ピボットテーブル',
        'アクティブなフィールド',
        'グループ',
        'フィルター',
        'データ',
        'アクション',
        '計算',
        'ツール',
        '表示',
      ],
      pivotDesign: [
        'レイアウト',
        'ピボットテーブル スタイルのオプション',
        'ピボットテーブル スタイル',
      ],
      sparklineDesign: ['スパークライン', '種類', '表示', 'スタイル', 'グループ'],
    },
    de: {
      tableDesign: [
        'Eigenschaften',
        'Tools',
        'Externe Tabellendaten',
        'Optionen für Tabellenformat',
        'Tabellenformatvorlagen',
      ],
      pivotAnalyze: [
        'PivotTable',
        'Aktives Feld',
        'Gruppe',
        'Filter',
        'Daten',
        'Aktionen',
        'Berechnungen',
        'Tools',
        'Einblenden',
      ],
      pivotDesign: ['Layout', 'Optionen für PivotTable-Formate', 'PivotTable-Formate'],
      sparklineDesign: ['Sparkline', 'Typ', 'Anzeigen', 'Format', 'Gruppe'],
    },
  },
  default: {
    en: {
      tableDesign: [
        'Properties',
        'Tools',
        'External Table Data',
        'Table Style Options',
        'Table Styles',
      ],
      pivotAnalyze: UNTITLED(9),
      pivotDesign: UNTITLED(4),
      sparklineDesign: ['Sparkline', 'Sparkline Type', 'Show', 'Sparkline Style', 'Group'],
    },
    de: {
      tableDesign: [
        'Eigenschaften',
        'Extras',
        'Externe Tabellendaten',
        'Optionen für Tabellenformatvorlagen',
        'Tabellenformatvorlagen',
      ],
      pivotAnalyze: UNTITLED(9),
      pivotDesign: UNTITLED(4),
      sparklineDesign: [
        'Sparkline',
        'Sparklinetyp',
        'Anzeigen',
        'Sparklineformatvorlage',
        'Gruppe',
      ],
    },
  },
};

/** Commands of the table Tools group and of each PivotTable Design group, in order. */
const LAYOUT: Record<Platform, { tableTools: string[]; pivotDesign: string[][] }> = {
  mac: {
    tableTools: [
      'tableSummarizePivot',
      'mac.data.removeDuplicates',
      'tableConvertRange',
      'mac.insert.slicer',
    ],
    pivotDesign: [
      ['pivotSubtotals', 'pivotGrandTotals', 'pivotReportLayout', 'pivotBlankRows'],
      ['pivotRowHeaders', 'pivotColHeaders', 'pivotRowStripes', 'pivotColStripes'],
      ['pivotQuickStyles'],
    ],
  },
  default: {
    tableTools: ['tableSummarizePivot', 'removeDupes', 'tableConvertRange', 'tableInsertSlicer'],
    pivotDesign: [
      ['pivotSubtotals', 'pivotGrandTotals'],
      ['pivotReportLayout', 'pivotBlankRows'],
      ['pivotRowHeaders', 'pivotColHeaders', 'pivotRowStripes', 'pivotColStripes'],
      ['pivotQuickStyles'],
    ],
  },
};

/** Each group of a tab's panel: its accessible name, its drawn caption, and its own commands
 *  (menu entries excluded), in ribbon order. */
const readGroups = (page: Page, tab: Tab) =>
  page.locator(`[data-ribbon-panel="${tab}"]`).evaluate((panel) =>
    Array.from(panel.querySelectorAll(':scope > section.fc-tb__ribbon-group')).map((group) => {
      const ids: string[] = [];
      for (const el of group.querySelectorAll<HTMLElement>('[data-ribbon-command]')) {
        if (el.closest('[role=menu]')) continue;
        const id = el.dataset.ribbonCommand ?? '';
        if (!ids.includes(id)) ids.push(id);
      }
      return {
        name: group.getAttribute('aria-label') ?? '',
        caption: group.querySelector(':scope > .fc-tb__ribbon-label')?.textContent ?? '',
        commands: ids,
      };
    }),
  );

/** The Table, PivotTable Analyze, PivotTable Design and Sparkline tabs title their groups in
 *  `locale`, the way `platform` draws them, and place each command in its group. */
export async function runContextualGroupTitlesScenario(
  page: Page,
  platform: Platform,
  locale: Locale,
): Promise<void> {
  const titles = TITLES[platform][locale];
  if (!titles) throw new Error(`no titles for ${platform}/${locale}`);
  const layout = LAYOUT[platform];
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform, locale });
  await page.setViewportSize({ width: 2560, height: 900 });
  await sp.expectNoStub();
  await page.evaluate(() => {
    const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
    const inst = w.__fcInst;
    const sheet = inst.workbook.sheet(0);
    inst.workbook.batch(() => {
      sheet.setText(0, 0, 'Channel');
      sheet.setText(0, 1, 'Sales');
      sheet.setText(1, 0, 'Online');
      sheet.setNumber(1, 1, 10);
      sheet.setText(2, 0, 'Retail');
      sheet.setNumber(2, 1, 20);
    });
    sheet.addTable({ range: 'A1:B3', name: 'Sales' });
    const handle = w.__fcProbes['api/workbook'].workbookHandleOf(inst.workbook);
    const made = w.__fcProbes['commands/pivot-table'].createPivotTableFromRange(handle, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
      destination: { sheet: 0, row: 0, col: 3 },
      name: 'ChannelReport',
      rowField: 'Channel',
      valueField: 'Sales',
      aggregation: 0,
    });
    if (!made.ok) throw new Error(`pivot creation failed: ${made.reason}`);
    sheet.addSparklineGroup({ type: 0, sparklines: [{ source: 'B2:B3', location: 'H2' }] });
    inst.history.clear();
  });

  const expectTitles = async (tab: Tab) => {
    await selectRibbonTab(page, tab);
    await expect(page.locator(`[data-ribbon-panel="${tab}"]`)).toBeVisible();
    const groups = await readGroups(page, tab);
    expect(groups.map((g) => g.name)).toEqual(titles[tab]);
    expect(groups.map((g) => g.caption)).toEqual(titles[tab]);
    return groups;
  };

  await sp.goTo('A2');
  const table = await expectTitles('tableDesign');
  expect(table[1]?.commands).toEqual(layout.tableTools);

  await sp.goTo('D2');
  await expectTitles('pivotAnalyze');
  const design = await expectTitles('pivotDesign');
  expect(design.map((g) => g.commands)).toEqual(layout.pivotDesign);

  await sp.goTo('H2');
  await expectTitles('sparklineDesign');
  await sp.expectNoConsoleErrors();
}
