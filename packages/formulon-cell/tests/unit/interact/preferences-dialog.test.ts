import { afterEach, describe, expect, it } from 'vitest';
import { spreadsheetFailureOf } from '../../../src/api/errors.js';
import { createDialogs } from '../../../src/api/spreadsheet/dialogs.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

const dialog = (): HTMLElement | null =>
  document.querySelector<HTMLElement>('.fc-prefs:not([hidden])');

const tool = (label: string): HTMLButtonElement | null =>
  dialog()?.querySelector<HTMLButtonElement>(`.fc-prefs__tool[aria-label="${label}"]`) ?? null;
/** Show All's visible groups, each with its visible pane buttons. */
const showAll = (): [string, string[]][] =>
  [...(dialog()?.querySelectorAll<HTMLElement>('.fc-prefs__group') ?? [])]
    .filter((group) => !group.hidden)
    .map((group) => [
      group.querySelector('.fc-prefs__group-title')?.textContent ?? '',
      [...group.querySelectorAll<HTMLButtonElement>('.fc-prefs__pane-button')]
        .filter((button) => !button.hidden)
        .map((button) => button.textContent ?? ''),
    ]);

describe('Preferences window', () => {
  let sheet: MountedStubSheet | undefined;
  afterEach(() => {
    sheet?.dispose();
    sheet = undefined;
    document.body.replaceChildren();
  });

  it('shows the zero values setting on the View pane, titled with the pane name', async () => {
    sheet = await mountStubSheet({
      locale: 'ja',
      ui: { platform: 'mac' },
      workbook: await WorkbookHandle.createDefault(),
    });
    const instance = sheet.instance;
    const dialogs = createDialogs(() => instance);
    expect(dialogs.canOpen('preferences')).toBe(true);
    dialogs.open('preferences');

    expect(dialog()?.getAttribute('aria-label')).toBe('表示');
    expect(dialog()?.querySelector('.fc-prefs__header')?.textContent).toBe('表示');
    expect(dialog()?.querySelector('.fc-prefs__heading')?.textContent).toBe('表示');
    const row = dialog()?.querySelector<HTMLLabelElement>('.fc-prefs__check');
    expect(row?.textContent).toBe('ゼロ表示');
    expect(row?.title).toBe('ゼロ値が格納されているセルに 0 を表示します。');
    const box = row?.querySelector('input') as HTMLInputElement;
    expect(box.checked).toBe(true);

    box.checked = false;
    box.dispatchEvent(new Event('change'));
    expect(instance.store.getState().ui.showZeros).toBe(false);
    expect(instance.workbook.getSheetView(0)?.showZeros).toBe(false);

    dialogs.close('preferences');
    expect(dialog()).toBeNull();
    dialogs.open('preferences');
    expect(dialog()?.querySelector<HTMLInputElement>('.fc-prefs__check input')?.checked).toBe(
      false,
    );
  });

  it.each([
    ['zh', '视图', '零值'],
    ['fr', 'Affichage', 'Valeurs nulles'],
  ] as const)(
    'titles the View pane as the desktop draws it in %s',
    async (locale, title, zeros) => {
      sheet = await mountStubSheet({ locale, ui: { platform: 'mac' } });
      const instance = sheet.instance;
      createDialogs(() => instance).open('preferences');
      expect(dialog()?.querySelector('.fc-prefs__header')?.textContent).toBe(title);
      expect(dialog()?.querySelector('.fc-prefs__check')?.textContent).toBe(zeros);
    },
  );

  it('carries the desktop toolbar: Back and Forward, Show All, the pane name and Search', async () => {
    sheet = await mountStubSheet({ locale: 'ja', ui: { platform: 'mac' } });
    const instance = sheet.instance;
    createDialogs(() => instance).open('preferences');
    const toolbar = dialog()?.querySelector<HTMLElement>('[role="toolbar"]');
    const labels = [...(toolbar?.querySelectorAll('button, input') ?? [])].map((control) =>
      control.getAttribute('aria-label'),
    );
    expect(labels).toEqual(['戻る', '進む', 'すべて表示', '検索']);
    expect(toolbar?.querySelector('[role="group"]')?.getAttribute('aria-label')).toBe('戻る/次へ');
    expect(tool('戻る')?.disabled).toBe(true);
    expect(tool('進む')?.disabled).toBe(true);

    tool('すべて表示')?.click();
    expect(dialog()?.getAttribute('aria-label')).toBe('環境設定');
    expect(dialog()?.querySelector('.fc-prefs__heading')?.textContent).toBe('');
    expect(showAll()).toEqual([
      ['オーサリング', ['表示', 'オートコレクト']],
      ['数式とリスト', ['計算']],
    ]);
    expect(dialog()?.querySelector<HTMLElement>('.fc-prefs__pane')?.hidden).toBe(true);
    expect(tool('戻る')?.disabled).toBe(false);
    expect(tool('進む')?.disabled).toBe(true);

    tool('戻る')?.click();
    expect(dialog()?.getAttribute('aria-label')).toBe('表示');
    expect(dialog()?.querySelector<HTMLElement>('.fc-prefs__all')?.hidden).toBe(true);
    expect(tool('戻る')?.disabled).toBe(true);
    expect(tool('進む')?.disabled).toBe(false);

    tool('進む')?.click();
    dialog()?.querySelector<HTMLButtonElement>('.fc-prefs__pane-button')?.click();
    expect(dialog()?.getAttribute('aria-label')).toBe('表示');
    expect(tool('進む')?.disabled).toBe(true);
  });

  it('narrows Show All to the panes whose names match the search', async () => {
    sheet = await mountStubSheet({ locale: 'en', ui: { platform: 'mac' } });
    const instance = sheet.instance;
    createDialogs(() => instance).open('preferences');
    const search = dialog()?.querySelector<HTMLInputElement>('input[type="search"]');
    const type = (value: string): void => {
      if (!search) return;
      search.value = value;
      search.dispatchEvent(new Event('input'));
    };

    type('vie');
    expect(dialog()?.getAttribute('aria-label')).toBe('Preferences');
    expect(showAll()).toEqual([['Authoring', ['View']]]);
    type('zzz');
    expect(showAll()).toEqual([]);
    type('calc');
    expect(showAll()).toEqual([['Formulas and Lists', ['Calculation']]]);
    type('');
    expect(showAll()).toEqual([
      ['Authoring', ['View', 'AutoCorrect']],
      ['Formulas and Lists', ['Calculation']],
    ]);
  });

  it('lists the AutoCorrect options in the desktop order, all on, sharing the dialog store', async () => {
    sheet = await mountStubSheet({ locale: 'en', ui: { platform: 'mac' } });
    const instance = sheet.instance;
    const dialogs = createDialogs(() => instance);
    dialogs.open('preferences');
    dialog()?.querySelector<HTMLButtonElement>('[data-pane="autoCorrect"]')?.click();
    expect(dialog()?.getAttribute('aria-label')).toBe('AutoCorrect');
    const pane = dialog()?.querySelector<HTMLElement>('.fc-prefs__pane[data-pane="autoCorrect"]');
    const rows = [...(pane?.querySelectorAll<HTMLLabelElement>('.fc-autocorrect__check') ?? [])];
    expect(rows.map((row) => row.textContent)).toEqual([
      'Replace internet and network paths with hyperlinks',
      'Capitalize names of days',
      'Correct TWo INitial CApitals',
      'Capitalize first letter of sentences',
      'Replace text as you type',
    ]);
    expect(rows.map((row) => row.querySelector('input')?.checked)).toEqual([
      true,
      true,
      true,
      true,
      true,
    ]);

    const days = rows[1]?.querySelector('input') as HTMLInputElement;
    days.checked = false;
    days.dispatchEvent(new Event('change'));
    expect(instance.autoCorrect.state().capitalizeDays).toBe(false);

    dialogs.close('preferences');
    dialogs.open('autoCorrect');
    const reopened = [
      ...document.querySelectorAll<HTMLLabelElement>('.fc-autocorrect .fc-autocorrect__check'),
    ].find((row) => row.textContent === 'Capitalize names of days');
    expect(reopened?.querySelector('input')?.checked).toBe(false);
    dialogs.close('autoCorrect');
  });

  it('keeps the pane and the AutoCorrect dialog ids apart while both are open', async () => {
    sheet = await mountStubSheet({ locale: 'en', ui: { platform: 'mac' } });
    const dialogs = createDialogs(() => sheet?.instance as NonNullable<typeof sheet>['instance']);
    dialogs.open('preferences');
    dialog()?.querySelector<HTMLButtonElement>('[data-pane="autoCorrect"]')?.click();
    dialogs.open('autoCorrect');
    const ids = [...document.querySelectorAll<HTMLElement>('[id]')].map((el) => el.id);
    expect(ids.filter((id, index) => ids.indexOf(id) !== index)).toEqual([]);
    dialogs.close('autoCorrect');
    dialogs.close('preferences');
  });

  const calcPane = (): HTMLElement | null =>
    dialog()?.querySelector<HTMLElement>('.fc-prefs__pane[data-pane="calculation"]') ?? null;
  const calcBox = (root: ParentNode | null, option: string): HTMLInputElement | null =>
    root?.querySelector<HTMLInputElement>(`[data-calc-option="${option}"]`) ?? null;

  it('shows the Mac calculation controls on the Calculation pane', async () => {
    sheet = await mountStubSheet({ locale: 'en', ui: { platform: 'mac' } });
    const instance = sheet.instance;
    createDialogs(() => instance).open('preferences');
    dialog()?.querySelector<HTMLButtonElement>('[data-pane="calculation"]')?.click();
    expect(dialog()?.getAttribute('aria-label')).toBe('Calculation');
    const pane = calcPane();
    expect(pane?.hidden).toBe(false);
    const labels = [...(pane?.querySelectorAll<HTMLElement>('.fc-iterdlg__row') ?? [])]
      .filter((row) => !row.closest('[hidden]'))
      .map((row) => row.textContent);
    expect(labels).toEqual([
      'Automatic',
      'Automatic except for data tables',
      'Manual',
      'Recalculate workbook before saving',
      'Use iterative calculation',
      'Maximum iterations:',
      'Maximum change:',
      expect.stringContaining('R1C1'),
      'Set precision as displayed',
      'Use 1904 date system',
    ]);
  });

  it('writes the same settings the Iterative Calculation dialog reads', async () => {
    sheet = await mountStubSheet({ locale: 'en', ui: { platform: 'mac' } });
    const instance = sheet.instance;
    const dialogs = createDialogs(() => instance);
    dialogs.open('preferences');
    dialog()?.querySelector<HTMLButtonElement>('[data-pane="calculation"]')?.click();
    const r1c1 = calcBox(calcPane(), 'r1c1') as HTMLInputElement;
    r1c1.checked = true;
    r1c1.dispatchEvent(new Event('change', { bubbles: true }));
    expect(instance.store.getState().ui.r1c1).toBe(true);

    dialogs.close('preferences');
    dialogs.open('iterative');
    expect(calcBox(document.querySelector('.fc-iterdlg'), 'r1c1')?.checked).toBe(true);
    dialogs.close('iterative');
  });

  it('keeps the pane and the dialog element ids and radio groups apart while both are open', async () => {
    sheet = await mountStubSheet({ locale: 'en', ui: { platform: 'mac' } });
    const dialogs = createDialogs(() => sheet?.instance as NonNullable<typeof sheet>['instance']);
    dialogs.open('preferences');
    dialog()?.querySelector<HTMLButtonElement>('[data-pane="calculation"]')?.click();
    dialogs.open('iterative');
    const ids = [...document.querySelectorAll<HTMLElement>('[id]')].map((el) => el.id);
    expect(ids.filter((id, index) => ids.indexOf(id) !== index)).toEqual([]);
    const groups = [
      ...document.querySelectorAll<HTMLInputElement>(
        '.fc-prefs [data-calc-option], .fc-iterdlg [data-calc-option]',
      ),
    ].map((input) => input.name);
    expect(new Set(groups.filter(Boolean)).size).toBe(2);
    dialogs.close('iterative');
    dialogs.close('preferences');
  });

  it('is refused on the default platform, which keeps the setting on its ribbon', async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    const instance = sheet.instance;
    const dialogs = createDialogs(() => instance);
    expect(dialogs.canOpen('preferences')).toBe(false);
    let failure: unknown;
    try {
      dialogs.open('preferences');
    } catch (error) {
      failure = error;
    }
    expect(spreadsheetFailureOf(failure)).toBe('unsupported-platform');
    expect(dialog()).toBeNull();
  });

  it('closes when the platform leaves the Mac profile', async () => {
    sheet = await mountStubSheet({ locale: 'en', ui: { platform: 'mac' } });
    const instance = sheet.instance;
    createDialogs(() => instance).open('preferences');
    expect(dialog()?.getAttribute('aria-label')).toBe('View');
    instance.setUi({ platform: 'default' });
    expect(dialog()).toBeNull();
  });
});
