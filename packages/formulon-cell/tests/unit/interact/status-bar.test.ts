import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { en, ja } from '../../../src/i18n/strings.js';
import { attachStatusBar } from '../../../src/interact/status-bar.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

const setRange = (
  store: SpreadsheetStore,
  r0: number,
  c0: number,
  r1: number,
  c1: number,
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      active: { sheet: 0, row: r0, col: c0 },
      anchor: { sheet: 0, row: r0, col: c0 },
      range: { sheet: 0, r0, c0, r1, c1 },
    },
  }));
};

const seedNumber = (store: SpreadsheetStore, row: number, col: number, value: number): void => {
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    cells.set(addrKey({ sheet: 0, row, col }), {
      value: { kind: 'number', value },
      formula: null,
    });
    return { ...s, data: { ...s.data, cells } };
  });
};

describe('attachStatusBar', () => {
  let statusbar: HTMLElement;
  let store: SpreadsheetStore;

  beforeEach(() => {
    statusbar = document.createElement('div');
    document.body.appendChild(statusbar);
    store = createSpreadsheetStore();
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('renders left/center/right segments without leaking the engine label by default', () => {
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
    });
    expect(statusbar.querySelector('.fc-host__statusbar-left')).not.toBeNull();
    expect(statusbar.querySelector('.fc-host__statusbar-aggs')).not.toBeNull();
    const right = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-right');
    expect(statusbar.querySelector('.fc-host__statusbar-left')?.textContent).toContain('準備完了');
    expect(right?.textContent).toContain('セル');
    expect(right?.textContent).not.toContain('stub');
    handle.detach();
  });

  it('can expose the engine label for dev/debug hosts', () => {
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
      showEngineLabel: true,
    });
    const right = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-right');
    expect(right?.textContent).toContain('stub');
    handle.detach();
  });

  it('reflects sum/avg/count for a numeric selection', () => {
    seedNumber(store, 0, 0, 10);
    seedNumber(store, 1, 0, 20);
    seedNumber(store, 2, 0, 30);
    setRange(store, 0, 0, 2, 0);

    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
    });
    const center = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-aggs');
    expect(center?.textContent).toContain('60'); // sum
    expect(center?.textContent).toContain('20'); // avg
    expect(center?.textContent).toContain('3'); // count
    handle.detach();
  });

  it('uses Excel default aggregate order: Average, Count, Sum', () => {
    seedNumber(store, 0, 0, 10);
    seedNumber(store, 1, 0, 20);
    setRange(store, 0, 0, 1, 0);

    const handle = attachStatusBar({
      statusbar,
      store,
      strings: en,
      getEngineLabel: () => 'stub',
    });
    const text =
      statusbar.querySelector<HTMLElement>('.fc-host__statusbar-aggs')?.textContent ?? '';
    const averageIndex = text.indexOf('Average');
    const countIndex = text.indexOf('Count');
    const sumIndex = text.indexOf('Sum');
    expect(averageIndex).toBeGreaterThanOrEqual(0);
    expect(countIndex).toBeGreaterThan(averageIndex);
    expect(sumIndex).toBeGreaterThan(countIndex);
    handle.detach();
  });

  it('reflects Excel-style Ready, Enter, Edit, and Point modes', () => {
    const handle = attachStatusBar({
      statusbar,
      store,
      strings: en,
      getEngineLabel: () => 'stub',
    });
    const left = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-left');
    expect(left?.textContent).toContain('Ready');

    mutators.setEditor(store, { kind: 'enter', raw: 'abc' });
    expect(left?.textContent).toContain('Enter');

    mutators.setEditor(store, { kind: 'edit', raw: 'abc', caret: 3 });
    expect(left?.textContent).toContain('Edit');

    mutators.setEditorRefs(store, [{ r0: 0, c0: 0, r1: 0, c1: 0, colorIndex: 0 }]);
    expect(left?.textContent).toContain('Point');
    handle.detach();
  });

  it('uses Edit mode while the formula bar is active', () => {
    let editing = false;
    const handle = attachStatusBar({
      statusbar,
      store,
      strings: en,
      getEngineLabel: () => 'stub',
      getFormulaEditing: () => editing,
    });
    const left = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-left');
    editing = true;
    handle.refresh();
    expect(left?.textContent).toContain('Edit');
    handle.detach();
  });

  it('keeps center empty when nothing is selected (and no aggs apply)', () => {
    setRange(store, 0, 0, 0, 0);
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
    });
    const center = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-aggs');
    expect(center?.textContent).toBe('');
    handle.detach();
  });

  it('right-click opens a chooser; clicking a row toggles the agg in the store', () => {
    seedNumber(store, 0, 0, 5);
    setRange(store, 0, 0, 0, 0);
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
    });

    statusbar.dispatchEvent(
      new MouseEvent('contextmenu', {
        bubbles: true,
        cancelable: true,
        clientX: 100,
        clientY: 200,
      }),
    );
    const chooser = document.querySelector<HTMLElement>('.fc-statusbar__chooser');
    expect(chooser).not.toBeNull();
    expect(chooser?.style.display).toBe('block');

    const items = chooser?.querySelectorAll<HTMLButtonElement>('.fc-statusbar__chooser-item');
    expect(items?.length).toBe(14);
    const headings = Array.from(
      chooser?.querySelectorAll<HTMLElement>('.fc-statusbar__chooser-heading') ?? [],
    ).map((heading) => heading.textContent);
    expect(headings).toEqual(['集計表示', 'ステータス バー項目']);
    expect(chooser?.querySelector('[role="separator"]')).toBeTruthy();
    expect(document.activeElement).toBe(items?.[0]);
    expect(items?.[0]?.getAttribute('role')).toBe('menuitemcheckbox');
    expect(items?.[0]?.getAttribute('aria-checked')).toBe('true');
    const firstCheck = items?.[0]?.querySelector<HTMLElement>('.fc-statusbar__chooser-check');
    expect(firstCheck?.textContent).toBe('');
    // Toggle "sum" off from the Excel default Average, Count, Sum set.
    const sumItem = Array.from(items ?? []).find((b) => b.textContent?.includes('合計'));
    expect(sumItem).toBeDefined();
    sumItem?.click();
    expect(store.getState().ui.statusAggs).not.toContain('sum');
    expect(sumItem?.getAttribute('aria-checked')).toBe('false');
    handle.detach();
  });

  it('keeps the status bar chooser close to Excel 365 desktop menu geometry', () => {
    const source = readFileSync(join(root, 'src/interact/status-bar.ts'), 'utf8');
    const css = readFileSync(
      join(root, 'src/styles/core/app/popups/validation-and-chooser.css'),
      'utf8',
    );
    const chooserCss = css.slice(css.indexOf('.fc-statusbar__chooser {'));

    expect(chooserCss).toMatch(
      /\.fc-statusbar__chooser\s*\{[\s\S]*?min-width: 232px;[\s\S]*?padding: 5px 0;[\s\S]*?border-radius: 2px;[\s\S]*?box-shadow:/,
    );
    expect(chooserCss).toMatch(
      /\.fc-statusbar__chooser-heading\s*\{[\s\S]*?padding: 5px 12px 3px 28px;[\s\S]*?text-transform: none;[\s\S]*?letter-spacing: 0;/,
    );
    expect(chooserCss).toMatch(
      /\.fc-statusbar__chooser-item\s*\{[\s\S]*?box-sizing: border-box;[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 4px;[\s\S]*?border-radius: 0;/,
    );
    expect(chooserCss).toMatch(
      /\.fc-statusbar__chooser-check\s*\{[\s\S]*?flex: 0 0 18px;[\s\S]*?width: 18px;[\s\S]*?height: 16px;/,
    );
    expect(chooserCss).toMatch(
      /\.fc-statusbar__chooser-item\[aria-checked="true"\] \.fc-statusbar__chooser-check::before\s*\{[\s\S]*?border-bottom: 1\.7px solid currentColor;[\s\S]*?content: "";/,
    );
    expect(chooserCss).not.toContain('text-transform: uppercase;');
    expect(chooserCss).not.toContain('background: var(--fc-accent-soft');
    expect(source).not.toContain("right.textContent = '—'");
  });

  it('positions the chooser above the opener and clamps it inside the viewport', () => {
    Object.defineProperty(window, 'innerWidth', { configurable: true, value: 320 });
    Object.defineProperty(window, 'innerHeight', { configurable: true, value: 180 });
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
    });
    const chooser = document.querySelector<HTMLElement>('.fc-statusbar__chooser');
    expect(chooser).toBeTruthy();
    if (chooser) {
      Object.defineProperty(chooser, 'offsetWidth', { configurable: true, value: 220 });
      Object.defineProperty(chooser, 'offsetHeight', { configurable: true, value: 100 });
    }

    statusbar.dispatchEvent(
      new MouseEvent('contextmenu', {
        bubbles: true,
        cancelable: true,
        clientX: 310,
        clientY: 170,
      }),
    );

    expect(chooser?.style.left).toBe('96px');
    expect(chooser?.style.top).toBe('62px');
    handle.detach();
  });

  it('right-click chooser toggles view shortcuts, zoom, and zoom slider', () => {
    const handle = attachStatusBar({
      statusbar,
      store,
      strings: en,
      getEngineLabel: () => 'stub',
    });
    statusbar.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
    const chooser = document.querySelector<HTMLElement>('.fc-statusbar__chooser');
    const rows = Array.from(
      chooser?.querySelectorAll<HTMLButtonElement>('.fc-statusbar__chooser-item') ?? [],
    );
    const viewShortcuts = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-views');
    const zoom = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-zoom');
    const zoomSlider = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-zoom-slider');

    rows.find((row) => row.textContent?.includes('View Shortcuts'))?.click();
    expect(store.getState().ui.statusOptions.viewShortcuts).toBe(false);
    expect(viewShortcuts?.style.display).toBe('none');

    rows.find((row) => row.textContent?.includes('Zoom Slider'))?.click();
    expect(store.getState().ui.statusOptions.zoomSlider).toBe(false);
    expect(zoomSlider?.style.display).toBe('none');

    rows.find((row) => row.textContent?.includes('Zoom'))?.click();
    expect(store.getState().ui.statusOptions.zoom).toBe(false);
    expect(zoom?.style.display).toBe('none');
    handle.detach();
  });

  it('can show host-driven upload status and macro recording indicators', () => {
    let upload: 'saved' | 'saving' | 'error' | null = 'saving';
    let recording = false;
    store.setState((s) => ({
      ...s,
      ui: {
        ...s.ui,
        statusOptions: { ...s.ui.statusOptions, uploadStatus: true, macroRecording: true },
      },
    }));
    const handle = attachStatusBar({
      statusbar,
      store,
      strings: en,
      getEngineLabel: () => 'stub',
      getUploadStatus: () => upload,
      getMacroRecording: () => recording,
    });

    const uploadEl = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-upload');
    const macroEl = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-macro');
    expect(uploadEl?.style.display).toBe('');
    expect(uploadEl?.textContent).toBe('Saving...');
    expect(uploadEl?.dataset.uploadStatus).toBe('saving');
    expect(macroEl?.style.display).toBe('');
    expect(macroEl?.textContent).toBe('Record Macro');

    upload = 'error';
    recording = true;
    handle.refresh();
    expect(uploadEl?.textContent).toBe('Upload failed');
    expect(uploadEl?.dataset.uploadStatus).toBe('error');
    expect(macroEl?.textContent).toBe('Recording Macro');
    expect(macroEl?.dataset.macroRecording).toBe('true');
    handle.detach();
  });

  it('localizes host-driven upload status and macro recording indicators in Japanese', () => {
    let upload: 'saved' | 'saving' | 'error' | null = 'saved';
    let recording = false;
    store.setState((s) => ({
      ...s,
      ui: {
        ...s.ui,
        statusOptions: { ...s.ui.statusOptions, uploadStatus: true, macroRecording: true },
      },
    }));
    const handle = attachStatusBar({
      statusbar,
      store,
      strings: ja,
      getEngineLabel: () => 'stub',
      getUploadStatus: () => upload,
      getMacroRecording: () => recording,
    });

    const uploadEl = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-upload');
    const macroEl = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-macro');
    expect(uploadEl?.textContent).toBe('保存済み');
    expect(macroEl?.textContent).toBe('マクロの記録');

    upload = 'saving';
    recording = true;
    handle.refresh();
    expect(uploadEl?.textContent).toBe('保存中...');
    expect(macroEl?.textContent).toBe('マクロ記録中');

    upload = 'error';
    handle.refresh();
    expect(uploadEl?.textContent).toBe('アップロード失敗');
    expect(uploadEl?.dataset.uploadStatus).toBe('error');
    expect(macroEl?.dataset.macroRecording).toBe('true');
    handle.detach();
  });

  it('keeps upload and macro indicators hidden without host drivers', () => {
    store.setState((s) => ({
      ...s,
      ui: {
        ...s.ui,
        statusOptions: { ...s.ui.statusOptions, uploadStatus: true, macroRecording: true },
      },
    }));
    const handle = attachStatusBar({
      statusbar,
      store,
      strings: en,
      getEngineLabel: () => 'stub',
    });

    expect(statusbar.querySelector<HTMLElement>('.fc-host__statusbar-upload')?.style.display).toBe(
      'none',
    );
    expect(statusbar.querySelector<HTMLElement>('.fc-host__statusbar-macro')?.style.display).toBe(
      'none',
    );
    handle.detach();
  });

  it('shows active keyboard lock indicators when their status options are enabled', () => {
    const handle = attachStatusBar({
      statusbar,
      store,
      strings: en,
      getEngineLabel: () => 'stub',
    });
    const locks = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-locks');
    expect(locks?.style.display).toBe('none');

    const capsEvent = new KeyboardEvent('keydown', { bubbles: true });
    Object.defineProperty(capsEvent, 'getModifierState', {
      value: (key: string) => key === 'CapsLock',
    });
    document.dispatchEvent(capsEvent);
    expect(locks?.textContent).toContain('Caps Lock');
    expect(locks?.style.display).toBe('');

    statusbar.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
    const chooser = document.querySelector<HTMLElement>('.fc-statusbar__chooser');
    const rows = Array.from(
      chooser?.querySelectorAll<HTMLButtonElement>('.fc-statusbar__chooser-item') ?? [],
    );
    rows.find((row) => row.textContent?.includes('Caps Lock'))?.click();
    expect(store.getState().ui.statusOptions.capsLock).toBe(false);
    expect(locks?.textContent).not.toContain('Caps Lock');
    expect(locks?.style.display).toBe('none');
    handle.detach();
  });

  it('Escape closes the chooser and restores focus to the opener', () => {
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
    });
    statusbar.tabIndex = -1;
    statusbar.focus();
    statusbar.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    const chooser = document.querySelector<HTMLElement>('.fc-statusbar__chooser');
    expect(chooser?.style.display).toBe('none');
    expect(document.activeElement).toBe(statusbar);
    handle.detach();
  });

  it('chooser supports arrow keys and Enter/Space toggles', () => {
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
    });
    statusbar.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
    const chooser = document.querySelector<HTMLElement>('.fc-statusbar__chooser');
    const items = chooser?.querySelectorAll<HTMLButtonElement>('.fc-statusbar__chooser-item');
    if (!items || items.length < 2) throw new Error('expected chooser items');

    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', cancelable: true }));
    expect(document.activeElement).toBe(items[items.length - 1]);
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', cancelable: true }));
    expect(document.activeElement).toBe(items[0]);
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', cancelable: true }));
    expect(document.activeElement).toBe(items[0]);
    document.dispatchEvent(new KeyboardEvent('keydown', { key: ' ', cancelable: true }));

    expect(store.getState().ui.statusAggs).not.toContain('average');
    expect(items[0]?.getAttribute('aria-checked')).toBe('false');
    handle.detach();
  });

  it('opens the chooser from keyboard context menu shortcuts', () => {
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
    });
    statusbar.tabIndex = -1;
    statusbar.focus();
    statusbar.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'F10', shiftKey: true, bubbles: true }),
    );
    let chooser = document.querySelector<HTMLElement>('.fc-statusbar__chooser');
    expect(chooser?.style.display).toBe('block');
    expect(document.activeElement).toBe(
      chooser?.querySelector<HTMLButtonElement>('.fc-statusbar__chooser-item'),
    );

    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    statusbar.dispatchEvent(new KeyboardEvent('keydown', { key: 'ContextMenu', bubbles: true }));
    chooser = document.querySelector<HTMLElement>('.fc-statusbar__chooser');
    expect(chooser?.style.display).toBe('block');
    handle.detach();
  });

  it('refresh() re-reads the engine label', () => {
    let label = 'stub';
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => label,
      showEngineLabel: true,
    });
    let right = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-right');
    expect(right?.textContent).toContain('stub');
    label = 'formulon 9.9.9';
    handle.refresh();
    right = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-right');
    expect(right?.textContent).toContain('formulon 9.9.9');
    handle.detach();
  });

  it('detach removes the chooser node and stops subscribing', () => {
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
    });
    handle.detach();
    expect(document.querySelector('.fc-statusbar__chooser')).toBeNull();

    // Mutating state after detach should not crash.
    mutators.setActive(store, { sheet: 0, row: 5, col: 5 });
  });

  it('hides the calc-mode badge when getCalcMode is omitted or returns null', () => {
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
      getCalcMode: () => null,
    });
    const badge = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-calcmode');
    expect(badge).not.toBeNull();
    expect(badge?.style.display).toBe('none');
    handle.detach();
  });

  it('renders the calc-mode badge with the active mode label', () => {
    let mode: 0 | 1 | 2 = 1; // Manual
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
      getCalcMode: () => mode,
    });
    const badge = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-calcmode');
    expect(badge?.style.display).toBe('');
    // defaultStrings is ja-JP; the test asserts the localized label.
    expect(badge?.textContent).toContain('手動');
    expect(badge?.dataset.calcMode).toBe('1');

    mode = 0;
    handle.refresh();
    expect(badge?.textContent).toContain('自動');
    expect(badge?.dataset.calcMode).toBe('0');
    handle.detach();
  });

  it('badge click invokes onCycleCalcMode; double-click invokes onRecalc', () => {
    const cycle: number[] = [];
    const recalcs: number[] = [];
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
      getCalcMode: () => 0,
      onCycleCalcMode: () => cycle.push(1),
      onRecalc: () => recalcs.push(1),
    });
    const badge = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-calcmode');
    badge?.click();
    badge?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, cancelable: true }));
    expect(cycle.length).toBe(1);
    expect(recalcs.length).toBe(1);
    handle.detach();
  });

  it('renders zoom controls and applies slider changes', () => {
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
    });
    const slider = statusbar.querySelector<HTMLInputElement>('.fc-host__statusbar-zoom-slider');
    const label = statusbar.querySelector<HTMLElement>('.fc-host__statusbar-zoom-label');
    expect(slider).not.toBeNull();
    expect(slider?.value).toBe('100');
    expect(label?.textContent).toBe('100%');
    expect(slider?.getAttribute('aria-label')).toBe('ズーム');

    if (!slider) throw new Error('expected zoom slider');
    slider.value = '150';
    slider.dispatchEvent(new Event('input'));
    expect(store.getState().viewport.zoom).toBe(1.5);
    expect(label?.textContent).toBe('150%');
    handle.detach();
  });

  it('projects disabled reasons when zoom buttons hit their limits', () => {
    mutators.setZoom(store, 0.5);
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
    });
    const buttons = statusbar.querySelectorAll<HTMLButtonElement>('.fc-host__statusbar-zoom-btn');
    const minus = buttons[0];
    const plus = buttons[1];
    expect(minus?.disabled).toBe(true);
    expect(minus?.dataset.disabledReason).toBe('ズーム倍率はすでに最小です。');
    expect(plus?.disabled).toBe(false);
    expect(plus?.dataset.disabledReason).toBeUndefined();

    mutators.setZoom(store, 4);
    handle.refresh();
    expect(minus?.disabled).toBe(false);
    expect(minus?.dataset.disabledReason).toBeUndefined();
    expect(plus?.disabled).toBe(true);
    expect(plus?.getAttribute('aria-description')).toBe('ズーム倍率はすでに最大です。');
    handle.detach();
  });

  it('renders workbook view shortcuts and switches the active view mode', () => {
    const handle = attachStatusBar({
      statusbar,
      store,
      strings: en,
      getEngineLabel: () => 'stub',
    });
    const buttons = statusbar.querySelectorAll<HTMLButtonElement>('.fc-host__statusbar-view');
    expect(buttons.length).toBe(3);
    expect(buttons[0]?.getAttribute('aria-label')).toBe('Normal');
    expect(buttons[0]?.getAttribute('aria-pressed')).toBe('true');
    expect(buttons[2]?.getAttribute('aria-label')).toBe('Page Break Preview');

    buttons[2]?.click();
    expect(store.getState().ui.workbookView).toBe('pageBreakPreview');
    expect(buttons[0]?.getAttribute('aria-pressed')).toBe('false');
    expect(buttons[2]?.getAttribute('aria-pressed')).toBe('true');
    handle.detach();
  });

  it('setStrings relabels static status and zoom chrome', () => {
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
    });
    handle.setStrings(en);
    expect(statusbar.querySelector('.fc-host__statusbar-left')?.textContent).toContain('Ready');
    expect(statusbar.querySelector('.fc-host__statusbar-right')?.textContent).toContain('cell');
    expect(
      statusbar
        .querySelector<HTMLInputElement>('.fc-host__statusbar-zoom-slider')
        ?.getAttribute('aria-label'),
    ).toBe('Zoom');
    handle.detach();
  });

  it('delegates zoom changes when onZoomChange is provided', () => {
    const calls: number[] = [];
    const handle = attachStatusBar({
      statusbar,
      store,
      getEngineLabel: () => 'stub',
      onZoomChange: (z) => {
        calls.push(z);
        mutators.setZoom(store, z);
      },
    });
    const plus = statusbar.querySelector<HTMLButtonElement>(
      '.fc-host__statusbar-zoom-btn:last-of-type',
    );
    plus?.click();
    expect(calls).toEqual([1.1]);
    expect(store.getState().viewport.zoom).toBeCloseTo(1.1);
    handle.detach();
  });

  it('keeps status bar button DOM on the shared interaction primitive', () => {
    const source = readFileSync(join(root, 'src/interact/status-bar.ts'), 'utf8');

    expect(source).toContain("import { createInteractionButton } from './chip-button.js'");
    expect(source).toContain('const createStatusBarButton');
    expect(source).toContain('const createStatusBarCalcButton');
    expect(source).toContain('const createStatusBarViewButton');
    expect(source).toContain('const createStatusBarZoomButton');
    expect(source).toContain('const createStatusBarChooserRow');
    expect(source).toContain('const calcBadge = createStatusBarCalcButton()');
    expect(source).toContain("normal: createStatusBarViewButton('normal'");
    expect(source).toContain("const zoomOut = createStatusBarZoomButton('−')");
    expect(source).toContain('const { row } = createStatusBarChooserRow(');
    expect(source).not.toContain("document.createElement('button')");
  });
});
