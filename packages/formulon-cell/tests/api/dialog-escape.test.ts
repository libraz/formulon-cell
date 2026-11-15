/**
 * Escape from a dialog's first control closes it and hands focus back to the sheet, for every
 * dialog `dialogs.open` reaches.
 */
import { ChartPlotType, PivotAggregation, PivotAxis } from '@libraz/formulon';
import { type DialogId, mount, type SpreadsheetInstance } from '@libraz/formulon-cell';
import { afterEach, describe, expect, it } from 'vitest';
import { expectFailure, makeHost } from './helpers.js';

type Platform = 'default' | 'mac';

// Every dialog id. A new id fails the type check until it is listed here.
const DIALOGS: Record<DialogId, true> = {
  conditional: true,
  iterative: true,
  externalLinks: true,
  cfRules: true,
  cellStylesGallery: true,
  evaluateFormula: true,
  functionArguments: true,
  hyperlink: true,
  comment: true,
  findReplace: true,
  pasteSpecial: true,
  insertCopiedCells: true,
  namedRange: true,
  defineName: true,
  pageSetup: true,
  format: true,
  dataValidation: true,
  goTo: true,
  goToSpecial: true,
  filterDropdown: true,
  watchWindow: true,
  quickAnalysis: true,
  workbookObjects: true,
  accessibilityChecker: true,
  chartSettings: true,
  pivotFieldList: true,
  activePivotFieldList: true,
  pivotFieldSettings: true,
  pivotOptions: true,
  pivotTable: true,
  spelling: true,
  autoCorrect: true,
  preferences: true,
};

// Ids the platform profile has no such dialog for: `dialogs.open` refuses them there.
const NOT_ON_PLATFORM: Record<Platform, readonly DialogId[]> = {
  default: ['preferences'],
  mac: ['quickAnalysis'],
};

// Docked, non-modal panes; their Escape contract is not a modal dialog's.
const PANES: Record<Platform, readonly DialogId[]> = {
  default: ['watchWindow', 'pivotFieldList', 'activePivotFieldList'],
  mac: [
    'watchWindow',
    'pivotFieldList',
    'activePivotFieldList',
    'accessibilityChecker',
    'externalLinks',
    'functionArguments',
  ],
};

// The cell each dialog opens on; the rest open on A1.
const ACTIVE_CELL: Partial<Record<DialogId, string>> = {
  pivotFieldList: 'D2',
  activePivotFieldList: 'D2',
  pivotFieldSettings: 'D2',
  pivotOptions: 'D2',
  spelling: 'H1',
};

const live: SpreadsheetInstance[] = [];
afterEach(() => {
  for (const d of live.splice(0).reverse()) d.dispose();
  document.body.replaceChildren();
});

const FOCUSABLE =
  'input:not([type="hidden"]):not([disabled]), select:not([disabled]), textarea:not([disabled]), button:not([disabled]), [tabindex]:not([tabindex="-1"])';

const shownSurfaces = (): HTMLElement[] =>
  Array.from(
    document.querySelectorAll<HTMLElement>(
      '[role="dialog"], [role="region"], [role="complementary"]',
    ),
  ).filter((el) => el.isConnected && !el.hidden && !el.closest('[hidden]'));

const pressEscape = (target: HTMLElement): void => {
  target.dispatchEvent(
    new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }),
  );
};

// A chart is selected by a key pressed on it, as by a click.
const selectChart = (host: HTMLElement): void => {
  const chart = host.ownerDocument.querySelector<HTMLElement>('.fc-chart');
  if (!chart) throw new Error('No chart.');
  chart.dispatchEvent(new KeyboardEvent('keydown', { key: 'Shift', bubbles: true }));
};

const mountWithContent = async (platform: Platform) => {
  const host = makeHost();
  const sheet = await mount(host, { locale: 'en', ui: { platform } });
  live.push(sheet);
  const s = sheet.workbook.sheet(0);
  s.setValues('A1', [
    ['Region', 'Amount'],
    ['East', 10],
    ['West', 20],
    ['East', 5],
  ]);
  s.setText('H1', 'Teh');
  const cache = sheet.workbook.addPivotCache();
  cache.setWorksheetSource({ range: 'A1:B4', sheet: s });
  cache.refresh();
  const pivot = s.addPivot({ name: 'P', cacheId: cache, anchor: 'D1' });
  pivot.setFieldAxis('Region', PivotAxis.Row);
  pivot.addDataField({ fieldIndex: 'Amount', aggregation: PivotAggregation.Sum });
  s.addChart({ type: ChartPlotType.Bar, range: 'B1:B4' });
  return { host, sheet };
};

describe('dialogs', () => {
  it.each(['default', 'mac'] as const)(
    'close on Escape from their first control and return focus to the sheet on the %s profile',
    async (platform) => {
      const { host, sheet } = await mountWithContent(platform);
      const failures: string[] = [];
      const walked: string[] = [];
      for (const id of Object.keys(DIALOGS) as DialogId[]) {
        if (NOT_ON_PLATFORM[platform].includes(id)) {
          if (sheet.dialogs.canOpen(id)) failures.push(`${id}: opens on this platform`);
          expectFailure(() => sheet.dialogs.open(id), 'unsupported-platform');
          continue;
        }
        if (PANES[platform].includes(id)) continue;
        sheet.selection.set(ACTIVE_CELL[id] ?? 'A1');
        if (id === 'pasteSpecial') sheet.commands.execute('copy');
        host.focus();
        if (id === 'chartSettings') selectChart(host);
        const before = new Set(shownSurfaces());
        sheet.dialogs.open(id);
        await Promise.resolve();
        const surface = shownSurfaces().find((el) => !before.has(el));
        if (!surface) {
          failures.push(`${id}: canOpen is true but nothing opened`);
          continue;
        }
        walked.push(id);
        const first = surface.querySelector<HTMLElement>(FOCUSABLE) ?? surface;
        first.focus();
        pressEscape(first);
        if (shownSurfaces().includes(surface)) {
          failures.push(`${id}: stayed open`);
          sheet.dialogs.close(id);
        } else if (!host.contains(document.activeElement)) {
          const active = document.activeElement;
          failures.push(`${id}: focus on ${active === document.body ? 'body' : active?.className}`);
        }
      }
      expect(failures).toEqual([]);
      expect(walked.length, walked.join(' ')).toBeGreaterThanOrEqual(24);
    },
    120_000,
  );

  it.each(['default', 'mac'] as const)(
    'leave docked panes open on Escape and hand focus to the sheet on the %s profile',
    async (platform) => {
      const { host, sheet } = await mountWithContent(platform);
      const failures: string[] = [];
      for (const id of PANES[platform]) {
        // Selecting a report cell may show the field list already, so look from before it.
        const before = new Set(shownSurfaces());
        sheet.selection.set(ACTIVE_CELL[id] ?? 'A1');
        host.focus();
        sheet.dialogs.open(id);
        await Promise.resolve();
        const pane = shownSurfaces().find((el) => !before.has(el));
        if (!pane) {
          failures.push(`${id}: nothing opened`);
          continue;
        }
        const first = pane.querySelector<HTMLElement>(FOCUSABLE) ?? pane;
        first.focus();
        pressEscape(first);
        if (!shownSurfaces().includes(pane)) failures.push(`${id}: closed on Escape`);
        else if (document.activeElement !== host) {
          failures.push(`${id}: focus on ${document.activeElement?.className}`);
        }
        sheet.dialogs.close(id);
      }
      expect(failures).toEqual([]);
    },
    120_000,
  );

  it.each(['default', 'mac'] as const)(
    'hand focus back to the button that opened them on the %s profile',
    async (platform) => {
      const { host } = await mountWithContent(platform);
      selectChart(host);
      const button = host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]');
      if (!button) throw new Error('No chart settings button.');
      button.focus();
      const before = new Set(shownSurfaces());
      button.click();
      const surface = shownSurfaces().find((el) => !before.has(el));
      if (!surface) throw new Error('Chart settings did not open.');
      pressEscape(surface.querySelector<HTMLElement>(FOCUSABLE) ?? surface);
      expect(shownSurfaces()).not.toContain(surface);
      expect(document.activeElement).toBe(button);
    },
  );
});
