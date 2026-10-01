import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import type { Range } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { defaultStrings } from '../../../src/i18n/strings.js';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';
import { toolbarMenuText } from '../../../src/toolbar/menu-text.js';
import { applyMergeAction } from '../../../src/toolbar/merge-action.js';
import { applyRibbonCommand } from '../../../src/toolbar/ribbon/apply-ribbon-command.js';
import { toolbarText } from '../../../src/toolbar/ribbon-model.js';

const makeWorkbook = (): WorkbookHandle =>
  ({
    capabilities: { merges: true },
    engineClearMerges: () => true,
    engineAddMerge: () => true,
    setBlank: () => undefined,
    setText: () => undefined,
  }) as unknown as WorkbookHandle;

const setRange = (store: SpreadsheetStore, range: Range): void => {
  store.setState((state) => ({
    ...state,
    selection: {
      active: { sheet: range.sheet, row: range.r0, col: range.c0 },
      anchor: { sheet: range.sheet, row: range.r0, col: range.c0 },
      range,
    },
  }));
};

const seedText = (store: SpreadsheetStore, row: number, col: number, value = 'x'): void => {
  store.setState((state) => {
    const cells = new Map(state.data.cells);
    cells.set(`${0}:${row}:${col}`, {
      value: { kind: 'text', value },
      formula: null,
    });
    return { ...state, data: { ...state.data, cells } };
  });
};

const contextFor = (store: SpreadsheetStore, history: History) => ({
  store,
  workbook: makeWorkbook(),
  history,
  strings: defaultStrings,
});

const ribbonInstanceFor = (store: SpreadsheetStore, history: History): SpreadsheetInstance =>
  ({
    store,
    history,
    workbook: makeWorkbook(),
    i18n: { strings: defaultStrings },
  }) as unknown as SpreadsheetInstance;

const clickButton = (label: string): void => {
  const button = [...document.body.querySelectorAll('button')].find(
    (candidate) => candidate.textContent === label,
  );
  button?.click();
};

describe('applyMergeAction', () => {
  it('merges across rows without a false data-loss warning', async () => {
    const store = createSpreadsheetStore();
    const history = new History();
    setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    seedText(store, 0, 0, 'top');
    seedText(store, 1, 0, 'bottom');

    expect(await applyMergeAction(contextFor(store, history), 'mergeAcross')).toBe(true);
    expect(document.body.querySelector('[role="alertdialog"]')).toBeNull();
    expect(store.getState().merges.byAnchor.size).toBe(2);
  });

  it('leaves the state unchanged when a data-loss warning is cancelled', async () => {
    const store = createSpreadsheetStore();
    const history = new History();
    setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    seedText(store, 0, 0, 'anchor');
    seedText(store, 0, 1, 'other');

    const pending = applyMergeAction(contextFor(store, history), 'mergeCells');
    await Promise.resolve();
    clickButton(defaultStrings.ribbon.mergeLoseDataCancel);

    expect(await pending).toBe(false);
    expect(store.getState().merges.byAnchor.size).toBe(0);
    expect(store.getState().data.cells.size).toBe(2);
  });

  it('records merge-and-center as one undoable transaction', async () => {
    const store = createSpreadsheetStore();
    const history = new History();
    setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });

    expect(await applyMergeAction(contextFor(store, history), 'mergeCenter')).toBe(true);
    expect(store.getState().merges.byAnchor.size).toBe(1);
    expect(store.getState().format.formats.get('0:0:0')?.align).toBe('center');

    expect(history.undo()).toBe(true);
    expect(store.getState().merges.byAnchor.size).toBe(0);
    expect(store.getState().format.formats.get('0:0:0')).toBeUndefined();
    expect(history.redo()).toBe(true);
    expect(store.getState().merges.byAnchor.size).toBe(1);
    expect(store.getState().format.formats.get('0:0:0')?.align).toBe('center');
  });

  it('does not align when the merge command rejects a one-cell range', async () => {
    const store = createSpreadsheetStore();
    const history = new History();
    setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });

    expect(await applyMergeAction(contextFor(store, history), 'mergeCenter')).toBe(false);
    expect(store.getState().format.formats.size).toBe(0);
  });

  it('uses Merge & Center for the main ribbon face', async () => {
    const store = createSpreadsheetStore();
    const history = new History();
    setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    const focused: boolean[] = [];
    const instance = ribbonInstanceFor(store, history);
    const handled = applyRibbonCommand('merge', {
      inst: instance,
      text: toolbarText(defaultStrings),
      menuText: toolbarMenuText(defaultStrings),
      ui: {
        theme: 'paper',
        borderStyle: 'thin',
        borderColor: '#000000',
        formulaBarVisible: true,
      },
      runtime: {
        focusSheet: () => focused.push(true),
        refreshCells: () => undefined,
        refreshZoom: () => undefined,
        projectFormatToolbar: () => undefined,
        applyRibbonFormat: () => undefined,
        applyUiTheme: () => undefined,
        setFormulaBarVisible: () => undefined,
        featureFlags: () => ({}),
        showMessage: () => undefined,
      },
    });

    expect(handled).toBe(true);
    expect(store.getState().merges.byAnchor.size).toBe(1);
    expect(store.getState().format.formats.get('0:0:0')?.align).toBe('center');
    await Promise.resolve();
    expect(focused).toEqual([true]);
  });
});
