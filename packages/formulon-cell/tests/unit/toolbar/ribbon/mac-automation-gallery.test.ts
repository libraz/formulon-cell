import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../../src/commands/history.js';
import { InteractionController } from '../../../../src/commands/interaction-controller.js';
import type {
  InteractionPolicy,
  OperationIntent,
  PermissionDecision,
} from '../../../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../../../src/commands/protection.js';
import type { CellValue, Range } from '../../../../src/engine/types.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { defaultStrings } from '../../../../src/i18n/strings.js';
import type { SpreadsheetInstance } from '../../../../src/mount/types.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import type { CellFormat } from '../../../../src/store/types.js';
import {
  disposeMacAutomationDialogs,
  MAC_AUTOMATION_GALLERY,
  openMacAutomationGallery,
  runMacAutomationScript,
} from '../../../../src/toolbar/ribbon/mac/automation-gallery.js';

interface FakeInstanceOptions {
  readonly selection?: Range;
  readonly getTables?: () => readonly {
    ref: string;
    sheetIndex: number;
    columns: readonly string[];
  }[];
  readonly getValue?: (addr: { sheet: number; row: number; col: number }) => CellValue;
  readonly canExecute?: (intent: OperationIntent) => PermissionDecision;
  readonly policy?: InteractionPolicy;
}

const liveHandles: WorkbookHandle[] = [];
const liveControllers: InteractionController[] = [];

class CountingFormatMap {
  private readonly values = new Map<string, CellFormat>();
  iterations = 0;

  get size(): number {
    return this.values.size;
  }

  set(key: string, value: CellFormat): this {
    this.values.set(key, value);
    return this;
  }

  get(key: string): CellFormat | undefined {
    return this.values.get(key);
  }

  *[Symbol.iterator](): Generator<[string, CellFormat]> {
    for (const entry of this.values) {
      this.iterations += 1;
      yield entry;
    }
  }
}

const makeInstance = (
  host: HTMLElement,
  options: FakeInstanceOptions = {},
): SpreadsheetInstance => {
  const store = createSpreadsheetStore();
  const workbook = {
    physicalCells: function* () {
      yield {
        addr: { sheet: 0, row: 0, col: 0 },
        value: { kind: 'text', value: 'header' },
        formula: null,
      };
      yield {
        addr: { sheet: 0, row: 2, col: 0 },
        value: { kind: 'number', value: 1 },
        formula: null,
      };
    },
    cells: function* () {
      yield* workbook.physicalCells(0);
    },
    getTables: options.getTables ?? (() => []),
    getValue: options.getValue ?? (() => ({ kind: 'text', value: 'value' })),
  } as unknown as WorkbookHandle;
  const instance = {
    host,
    store,
    workbook,
    i18n: { locale: 'ja-JP' },
    commands: {
      policy: options.policy,
      canExecute: options.canExecute ?? (() => ({ allowed: true })),
    },
  } as unknown as SpreadsheetInstance;
  if (options.selection) mutators.setRange(store, options.selection);
  return instance;
};

const makeNativeInstance = async (
  host: HTMLElement,
  options: { readonly getBounds?: () => Range } = {},
): Promise<{
  instance: SpreadsheetInstance;
  workbook: WorkbookHandle;
  store: SpreadsheetStore;
  history: History;
  controller: InteractionController;
}> => {
  const workbook = await WorkbookHandle.createDefault();
  liveHandles.push(workbook);
  const store = createSpreadsheetStore();
  const history = new History();
  const controller = new InteractionController({
    store,
    getWb: () => workbook,
    history,
    getBounds: options.getBounds,
  });
  liveControllers.push(controller);
  workbook.attachHistory(history);
  const instance = {
    host,
    store,
    workbook,
    history,
    commands: controller,
    i18n: { locale: 'ja-JP' },
  } as unknown as SpreadsheetInstance;
  return { instance, workbook, store, history, controller };
};

describe('Mac automation gallery', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
  });

  afterEach(() => {
    for (const controller of liveControllers.splice(0)) controller.dispose();
    for (const workbook of liveHandles.splice(0)) workbook.dispose();
    document.body.innerHTML = '';
  });

  it('ships bilingual descriptors for every local sample', () => {
    expect(
      MAC_AUTOMATION_GALLERY.every((entry) => entry.label.length > 0 && entry.labelJa.length > 0),
    ).toBe(true);
    expect(MAC_AUTOMATION_GALLERY.map((entry) => entry.id)).toEqual([
      'allRowsColumns',
      'freezeSelection',
      'makeSubtable',
      'removeHyperlinks',
      'countEmptyRows',
      'tableToJson',
      'newPivotTable',
    ]);
  });

  it('shows a visible Japanese result when counting blank rows', async () => {
    const instance = makeInstance(host);
    const pending = runMacAutomationScript(instance, 'countEmptyRows');
    const dialog = document.querySelector<HTMLElement>('.fc-macautomationresult');
    expect(dialog?.hidden).toBe(false);
    expect(dialog?.textContent).toContain('空白行: 1 行');
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await pending;
    expect(host.dataset.fcMacAutomationResult).toContain('blankRows');
  });

  it('localizes the gallery action labels from the active controller locale', async () => {
    const instance = makeInstance(host);
    const pending = openMacAutomationGallery(instance);
    const first = document.querySelector<HTMLButtonElement>('[data-mac-automation-script]');
    const cancel = document.querySelector<HTMLButtonElement>(
      '.fc-macautomationdlg .fc-fmtdlg__btn',
    );
    expect(first?.textContent).toBe('すべての行と列を表示');
    expect(cancel?.textContent).toBe('キャンセル');
    cancel?.click();
    await pending;
  });

  it('keeps JSON visible when clipboard access is unavailable', async () => {
    const instance = makeInstance(host);
    const clipboard = navigator.clipboard;
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: undefined,
    });
    const pending = runMacAutomationScript(instance, 'tableToJson');
    const output = document.querySelector<HTMLTextAreaElement>(
      '[data-mac-automation-result="true"]',
    );
    expect(output?.value).toContain('value');
    expect(instance.host.dataset.fcMacAutomationResult).toContain('copied');
    output
      ?.closest<HTMLElement>('.fc-macautomationresult')
      ?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')
      ?.click();
    await pending;
    Object.defineProperty(navigator, 'clipboard', { configurable: true, value: clipboard });
  });

  it('rejects a full-grid JSON selection before reading cells or emitting output', async () => {
    let reads = 0;
    const instance = makeInstance(host, {
      selection: { sheet: 0, r0: 0, c0: 0, r1: 1_048_575, c1: 16_383 },
      getValue: () => {
        reads += 1;
        return { kind: 'number', value: reads };
      },
    });

    await runMacAutomationScript(instance, 'tableToJson');

    expect(reads).toBe(0);
    expect(host.dataset.fcMacAutomationResult).toBeUndefined();
    expect(document.querySelector('.fc-macautomationresult')).toBeNull();
  });

  it('rejects an oversized table before reading cells or emitting output', async () => {
    let reads = 0;
    const instance = makeInstance(host, {
      selection: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
      getTables: () => [{ ref: 'A1:Z10000', sheetIndex: 0, columns: [] }],
      getValue: () => {
        reads += 1;
        return { kind: 'number', value: reads };
      },
    });

    await runMacAutomationScript(instance, 'tableToJson');

    expect(reads).toBe(0);
    expect(host.dataset.fcMacAutomationResult).toBeUndefined();
    expect(document.querySelector('.fc-macautomationresult')).toBeNull();
  });

  it('denies JSON before the first cell read and every visible side effect', async () => {
    let reads = 0;
    const instance = makeInstance(host, {
      policy: { defaultOperation: 'deny' },
      canExecute: () => ({ allowed: false, code: 'operationDenied', reason: 'export denied' }),
      getValue: () => {
        reads += 1;
        return { kind: 'number', value: reads };
      },
      selection: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
    });

    await runMacAutomationScript(instance, 'tableToJson');

    expect(reads).toBe(0);
    expect(host.dataset.fcMacAutomationResult).toBeUndefined();
    expect(document.querySelector('.fc-macautomationresult')).toBeNull();
  });

  it('authorizes JSON before the first native cell read', async () => {
    const events: string[] = [];
    const instance = makeInstance(host, {
      selection: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
      canExecute: () => {
        events.push('authorize');
        return { allowed: true };
      },
      getValue: () => {
        events.push('read');
        return { kind: 'number', value: 1 };
      },
    });

    await runMacAutomationScript(instance, 'tableToJson');

    expect(events[0]).toBe('authorize');
    expect(events).toContain('read');
    expect(host.dataset.fcMacAutomationResult).toContain('json');
    document
      .querySelector<HTMLElement>('.fc-macautomationresult .fc-fmtdlg__btn--primary')
      ?.click();
  });

  it('freezes, undoes, and redoes through one restricted policy transaction', async () => {
    const native = await makeNativeInstance(host);
    const { instance, workbook, store, history } = native;
    mutators.setRange(store, { sheet: 0, r0: 2, c0: 1, r1: 4, c1: 3 });
    expect(workbook.isStub).toBe(false);
    expect(workbook.capabilities.freeze).toBe(true);
    const origins: string[] = [];
    native.controller.setPolicy({
      defaultOperation: 'deny',
      operations: { format: true },
      restrict: ({ intent }) => {
        origins.push(intent.origin);
        return true;
      },
    });

    await runMacAutomationScript(instance, 'freezeSelection');

    expect(store.getState().layout.freezeRows).toBe(2);
    expect(store.getState().layout.freezeCols).toBe(1);
    expect(workbook.getSheetView(0)?.freezeRows).toBe(2);
    expect(workbook.getSheetView(0)?.freezeCols).toBe(1);
    expect(history.canUndo()).toBe(true);
    expect(origins).toEqual(['ribbon', 'undo', 'redo']);

    expect(history.undo()).toBe(true);
    expect(store.getState().layout.freezeRows).toBe(0);
    expect(store.getState().layout.freezeCols).toBe(0);
    expect(workbook.getSheetView(0)?.freezeRows).toBe(0);
    expect(workbook.getSheetView(0)?.freezeCols).toBe(0);

    expect(history.redo()).toBe(true);
    expect(store.getState().layout.freezeRows).toBe(2);
    expect(store.getState().layout.freezeCols).toBe(1);
    expect(workbook.getSheetView(0)?.freezeRows).toBe(2);
    expect(workbook.getSheetView(0)?.freezeCols).toBe(1);
  });

  it('denies an inverse Freeze policy before mutation or history creation', async () => {
    const native = await makeNativeInstance(host);
    const { instance, workbook, store, history } = native;
    mutators.setRange(store, { sheet: 0, r0: 2, c0: 1, r1: 4, c1: 3 });
    native.controller.setPolicy({
      defaultOperation: 'deny',
      operations: { format: true },
      restrict: ({ intent }) => intent.origin !== 'undo',
    });

    await runMacAutomationScript(instance, 'freezeSelection');

    expect(store.getState().layout.freezeRows).toBe(0);
    expect(store.getState().layout.freezeCols).toBe(0);
    expect(workbook.getSheetView(0)?.freezeRows).toBe(0);
    expect(workbook.getSheetView(0)?.freezeCols).toBe(0);
    expect(history.canUndo()).toBe(false);
    expect(host.dataset.fcMacAutomationResult).toBeUndefined();
  });

  it('reports Freeze rollback failures with both original and rollback errors', async () => {
    const native = await makeNativeInstance(host);
    const { instance, workbook, store, history, controller } = native;
    mutators.setRange(store, { sheet: 0, r0: 2, c0: 1, r1: 4, c1: 3 });
    controller.setPolicy({ defaultOperation: 'deny', operations: { format: true } });
    workbook.setSheetFreeze = () => {
      throw new Error('native freeze failed');
    };
    history.abort = () => {
      throw new Error('history rollback failed');
    };

    let caught: unknown;
    try {
      await runMacAutomationScript(instance, 'freezeSelection');
    } catch (error) {
      caught = error;
    }

    expect(caught).toBeInstanceOf(AggregateError);
    const aggregate = caught as AggregateError;
    expect(aggregate.errors).toHaveLength(2);
    expect((aggregate.errors[0] as Error).message).toBe('native freeze failed');
    expect((aggregate.errors[1] as Error).message).toBe('history rollback failed');
    expect(host.dataset.fcMacAutomationResult).toBeUndefined();
  });

  it('authorizes show-all-rows-and-columns replay for both resize operations', async () => {
    const { instance, history } = await makeNativeInstance(host);
    const seen = new Set<string>();
    (instance as unknown as { commands: unknown }).commands = {
      canExecute: (intent: OperationIntent): PermissionDecision => {
        seen.add(`${intent.operation}:${intent.origin}`);
        return { allowed: true };
      },
    };

    await runMacAutomationScript(instance, 'allRowsColumns');

    expect([...seen].sort()).toEqual([
      'resizeColumns:redo',
      'resizeColumns:ribbon',
      'resizeColumns:undo',
      'resizeRows:redo',
      'resizeRows:ribbon',
      'resizeRows:undo',
    ]);
    expect(history.canUndo()).toBe(false);
  });

  it('reports a failing gallery script instead of leaving an unhandled rejection', async () => {
    const instance = makeInstance(host);
    (instance as unknown as { i18n: unknown }).i18n = { locale: 'en-US', strings: defaultStrings };
    (instance as unknown as { openPivotTableDialog: () => void }).openPivotTableDialog = () => {
      throw new Error('pivot exploded');
    };
    const pending = openMacAutomationGallery(instance);
    document
      .querySelector<HTMLButtonElement>('[data-mac-automation-script="newPivotTable"]')
      ?.click();
    await pending;
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(document.body.textContent).toContain('pivot exploded');
  });

  it('removes and restores a native internal hyperlink losslessly', async () => {
    const native = await makeNativeInstance(host);
    const { instance, workbook, store, history } = native;
    expect(workbook.isStub).toBe(false);
    expect(workbook.capabilities.hyperlinks).toBe(true);
    expect(workbook.supportsHyperlinkRangeWrite()).toBe(true);
    expect(workbook.addHyperlinkRange(0, 40, 2, 41, 3, '', 'Jump', 'Tip', 'Sheet1!A1')).toBe(true);
    expect(
      workbook.addHyperlinkRange(0, 40, 2, 40, 2, 'https://example.com', 'Alt', 'Alt tip', ''),
    ).toBe(true);
    mutators.setCellFormat(store, { sheet: 0, row: 40, col: 2 }, { color: '#123456' });
    const before = workbook.getHyperlinksFull(0);
    expect(before).toEqual([
      {
        row: 40,
        col: 2,
        lastRow: 41,
        lastCol: 3,
        target: '',
        location: 'Sheet1!A1',
        display: 'Jump',
        tooltip: 'Tip',
      },
      {
        row: 40,
        col: 2,
        lastRow: 40,
        lastCol: 2,
        target: 'https://example.com',
        location: '',
        display: 'Alt',
        tooltip: 'Alt tip',
      },
    ]);
    expect(workbook.getHyperlinks(0)).toEqual([
      { row: 40, col: 2, target: '', display: 'Jump', tooltip: 'Tip' },
      { row: 40, col: 2, target: 'https://example.com', display: 'Alt', tooltip: 'Alt tip' },
    ]);
    const origins: string[] = [];
    native.controller.setPolicy({
      defaultOperation: 'deny',
      operations: { hyperlink: true },
      editable: ({ addr }) => addr.row >= 40 && addr.row <= 41 && addr.col >= 2 && addr.col <= 3,
      restrict: ({ intent }) => {
        origins.push(intent.origin);
        return true;
      },
    });

    await runMacAutomationScript(instance, 'removeHyperlinks');

    expect(workbook.getHyperlinksFull(0)).toEqual([]);
    expect(store.getState().format.formats.get('0:40:2')?.color).toBe('#123456');
    expect(host.dataset.fcMacAutomationResult).toContain('removed');
    expect(history.canUndo()).toBe(true);
    expect(origins).toEqual(['ribbon', 'undo', 'redo']);

    expect(history.undo()).toBe(true);
    expect(workbook.getHyperlinksFull(0)).toEqual(before);
    expect(store.getState().format.formats.get('0:40:2')?.color).toBe('#123456');
    const bytes = workbook.save();
    const loaded = await WorkbookHandle.loadBytes(bytes);
    liveHandles.push(loaded);
    expect(loaded.getHyperlinksFull(0)).toEqual(before);

    expect(history.redo()).toBe(true);
    expect(workbook.getHyperlinksFull(0)).toEqual([]);
    expect(store.getState().format.formats.get('0:40:2')?.hyperlink).toBeUndefined();
  });

  it('scopes hyperlink removal to the active sheet across native and store links', async () => {
    const native = await makeNativeInstance(host);
    const { instance, workbook, store, history, controller } = native;
    const otherSheet = workbook.addSheet('Other');
    expect(otherSheet).toBe(1);
    expect(workbook.addHyperlinkRange(0, 82, 1, 82, 1, 'https://active.example', 'Active')).toBe(
      true,
    );
    expect(
      workbook.addHyperlinkRange(1, 2, 2, 2, 2, 'https://other-native.example', 'Other native'),
    ).toBe(true);
    const otherStoreKey = '1:3:2';
    const scopedFormats = new Map<string, CellFormat>();
    for (let row = 0; row < 100_001; row += 1)
      scopedFormats.set(`1:${row}:0`, { hyperlink: `https://other-store-${row}.example` });
    scopedFormats.set(otherStoreKey, {
      hyperlink: 'https://other-store.example',
      hyperlinkDisplay: 'Other store',
    });
    scopedFormats.set('0:83:1', {
      hyperlink: 'https://active-store.example',
      hyperlinkDisplay: 'Active store',
    });
    store.setState((state) => ({
      ...state,
      format: { ...state.format, formats: scopedFormats },
    }));
    const activeBefore = workbook.getHyperlinksFull(0);
    const otherBefore = workbook.getHyperlinksFull(1);
    controller.setPolicy({
      defaultOperation: 'deny',
      operations: { hyperlink: true },
      editable: ({ addr }) => addr.sheet === 0 && addr.row >= 82 && addr.row <= 83,
    });

    await runMacAutomationScript(instance, 'removeHyperlinks');

    expect(workbook.getHyperlinksFull(0)).toEqual([]);
    expect(workbook.getHyperlinksFull(1)).toEqual(otherBefore);
    expect(store.getState().format.formats.get('0:83:1')?.hyperlink).toBeUndefined();
    expect(store.getState().format.formats.get(otherStoreKey)?.hyperlink).toBe(
      'https://other-store.example',
    );
    expect(history.canUndo()).toBe(true);

    expect(history.undo()).toBe(true);
    expect(workbook.getHyperlinksFull(0)).toEqual(activeBefore);
    expect(workbook.getHyperlinksFull(1)).toEqual(otherBefore);
    expect(store.getState().format.formats.get('0:83:1')?.hyperlink).toBe(
      'https://active-store.example',
    );
    expect(store.getState().format.formats.get(otherStoreKey)?.hyperlink).toBe(
      'https://other-store.example',
    );

    expect(history.redo()).toBe(true);
    expect(workbook.getHyperlinksFull(0)).toEqual([]);
    expect(workbook.getHyperlinksFull(1)).toEqual(otherBefore);
    expect(store.getState().format.formats.get('0:83:1')?.hyperlink).toBeUndefined();
    expect(store.getState().format.formats.get(otherStoreKey)?.hyperlink).toBe(
      'https://other-store.example',
    );
  });

  it('stops store hyperlink discovery at the materialization cap', async () => {
    const native = await makeNativeInstance(host);
    const { instance, store, history, controller, workbook } = native;
    const formats = new CountingFormatMap();
    for (let row = 0; row < 100_002; row += 1)
      formats.set(`0:${row}:0`, { hyperlink: `https://store-${row}.example` });
    store.setState((state) => ({
      ...state,
      format: { ...state.format, formats: formats as unknown as Map<string, CellFormat> },
    }));
    let authorizationCalls = 0;
    const canExecute = controller.canExecute.bind(controller);
    controller.canExecute = (intent) => {
      authorizationCalls += 1;
      return canExecute(intent);
    };
    controller.setPolicy({ defaultOperation: 'deny', operations: { hyperlink: true } });
    let events = 0;
    host.addEventListener('fc:mac-automation', () => {
      events += 1;
    });

    await runMacAutomationScript(instance, 'removeHyperlinks');

    expect(formats.iterations).toBe(100_001);
    expect(authorizationCalls).toBe(0);
    expect(workbook.getHyperlinksFull(0)).toEqual([]);
    expect(history.canUndo()).toBe(false);
    expect(events).toBe(0);
    expect(host.dataset.fcMacAutomationResult).toBeUndefined();
    expect(store.getState().format.formats.size).toBe(100_002);
  });

  it('stops a mixed native and store hyperlink union before the cap overflow', async () => {
    const native = await makeNativeInstance(host);
    const { instance, store, history, controller, workbook } = native;
    expect(workbook.addHyperlinkRange(0, 0, 0, 0, 0, 'https://native.example', 'Native')).toBe(
      true,
    );
    const formats = new CountingFormatMap();
    for (let row = 1; row <= 100_001; row += 1)
      formats.set(`0:${row}:0`, { hyperlink: `https://store-${row}.example` });
    store.setState((state) => ({
      ...state,
      format: { ...state.format, formats: formats as unknown as Map<string, CellFormat> },
    }));
    let authorizationCalls = 0;
    const canExecute = controller.canExecute.bind(controller);
    controller.canExecute = (intent) => {
      authorizationCalls += 1;
      return canExecute(intent);
    };
    controller.setPolicy({ defaultOperation: 'deny', operations: { hyperlink: true } });

    await runMacAutomationScript(instance, 'removeHyperlinks');

    expect(formats.iterations).toBe(100_000);
    expect(authorizationCalls).toBe(0);
    expect(workbook.getHyperlinksFull(0)).toHaveLength(1);
    expect(history.canUndo()).toBe(false);
    expect(host.dataset.fcMacAutomationResult).toBeUndefined();
    expect(store.getState().format.formats.size).toBe(100_001);
  });

  it('denies hyperlink removal before native mutation for bounds, editable, and protection', async () => {
    const native = await makeNativeInstance(host);
    const { instance, workbook, store, history } = native;
    expect(
      workbook.addHyperlinkRange(0, 60, 5, 60, 5, 'https://example.com', 'Link', 'Tip', ''),
    ).toBe(true);
    const before = workbook.getHyperlinksFull(0);
    let removes = 0;
    let adds = 0;
    const remove = workbook.removeHyperlink.bind(workbook);
    const add = workbook.addHyperlinkRange.bind(workbook);
    workbook.removeHyperlink = (...args) => {
      removes += 1;
      return remove(...args);
    };
    workbook.addHyperlinkRange = (...args) => {
      adds += 1;
      return add(...args);
    };
    native.controller.setPolicy({
      defaultOperation: 'deny',
      operations: { hyperlink: true },
      editable: { ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }] },
    });

    await runMacAutomationScript(instance, 'removeHyperlinks');

    expect(workbook.getHyperlinksFull(0)).toEqual(before);
    expect(removes).toBe(0);
    expect(adds).toBe(0);
    expect(history.canUndo()).toBe(false);
    expect(host.dataset.fcMacAutomationResult).toBeUndefined();

    native.controller.setPolicy({
      defaultOperation: 'deny',
      operations: { hyperlink: true },
      restrict: () => true,
    });
    setProtectedSheet(store, 0, true);
    await runMacAutomationScript(instance, 'removeHyperlinks');
    expect(workbook.getHyperlinksFull(0)).toEqual(before);
    expect(removes).toBe(0);
    expect(adds).toBe(0);

    const bounded = await makeNativeInstance(host, {
      getBounds: () => ({ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }),
    });
    expect(
      bounded.workbook.addHyperlinkRange(
        0,
        61,
        5,
        61,
        5,
        'https://bounds.example',
        'Bounds',
        '',
        '',
      ),
    ).toBe(true);
    const boundedBefore = bounded.workbook.getHyperlinksFull(0);
    bounded.controller.setPolicy({ defaultOperation: 'deny', operations: { hyperlink: true } });
    await runMacAutomationScript(bounded.instance, 'removeHyperlinks');
    expect(bounded.workbook.getHyperlinksFull(0)).toEqual(boundedBefore);
    expect(bounded.history.canUndo()).toBe(false);

    expect(
      native.controller.canExecute({
        operation: 'hyperlink',
        origin: 'ribbon',
        commandId: 'other.hyperlink.command',
        effects: [{ kind: 'cells', cells: [{ sheet: 0, row: 60, col: 5 }] }],
      }),
    ).toMatchObject({ allowed: false, code: 'unsupported' });
  });

  it('replays singular hyperlink history through policy and legacy guards', async () => {
    const native = await makeNativeInstance(host);
    const { controller, history, store } = native;
    const addr = { sheet: 0, row: 80, col: 4 };
    const cells = [addr];
    const makeIntent = (origin: 'ribbon' | 'undo' | 'redo') => ({
      operation: 'hyperlink' as const,
      origin,
      commandId: 'mac.automate.removeHyperlinks',
      effects: [{ kind: 'cells' as const, cells }],
    });
    const callbacks: string[] = [];
    const origins: string[] = [];
    controller.setPolicy({
      defaultOperation: 'deny',
      operations: { hyperlink: true },
      editable: ({ addr: candidate }) => candidate.row === addr.row && candidate.col === addr.col,
      restrict: ({ intent }) => {
        origins.push(intent.origin);
        return true;
      },
    });
    const undoIntent = makeIntent('undo');
    const redoIntent = makeIntent('redo');
    history.push({
      undo: () => callbacks.push('undo'),
      redo: () => callbacks.push('redo'),
      intent: redoIntent,
      inverseIntent: undoIntent,
    });

    expect(history.undo()).toBe(true);
    expect(history.redo()).toBe(true);
    expect(callbacks).toEqual(['undo', 'redo']);
    expect(origins).toEqual(['undo', 'redo']);

    controller.setPolicy(undefined);
    setProtectedSheet(store, 0, true);
    expect(history.undo()).toBe(false);
    expect(callbacks).toEqual(['undo', 'redo']);
    setProtectedSheet(store, 0, false);
    expect(history.undo()).toBe(true);
    expect(callbacks).toEqual(['undo', 'redo', 'undo']);

    const invalidAddr = { sheet: 0, row: 1_048_576, col: 4 };
    const invalidCells = [invalidAddr];
    history.push({
      undo: () => callbacks.push('invalid-undo'),
      redo: () => callbacks.push('invalid-redo'),
      intent: { ...makeIntent('redo'), effects: [{ kind: 'cells', cells: invalidCells }] },
      inverseIntent: { ...makeIntent('undo'), effects: [{ kind: 'cells', cells: invalidCells }] },
    });
    expect(history.undo()).toBe(false);
    expect(callbacks).toEqual(['undo', 'redo', 'undo']);
  });

  it('rolls back a partial native hyperlink removal and reports the failure without history or result', async () => {
    const native = await makeNativeInstance(host);
    const { instance, workbook, history } = native;
    expect(workbook.addHyperlinkRange(0, 70, 1, 70, 1, 'https://one.example', 'One', '', '')).toBe(
      true,
    );
    expect(workbook.addHyperlinkRange(0, 72, 1, 72, 1, 'https://two.example', 'Two', '', '')).toBe(
      true,
    );
    const before = workbook.getHyperlinksFull(0);
    let calls = 0;
    const remove = workbook.removeHyperlink.bind(workbook);
    workbook.removeHyperlink = (...args) => {
      calls += 1;
      if (calls === 2) return false;
      return remove(...args);
    };
    native.controller.setPolicy({ defaultOperation: 'deny', operations: { hyperlink: true } });

    await expect(runMacAutomationScript(instance, 'removeHyperlinks')).rejects.toThrow(
      'Could not remove hyperlink',
    );

    expect(calls).toBeGreaterThanOrEqual(2);
    expect(workbook.getHyperlinksFull(0)).toEqual(before);
    expect(history.canUndo()).toBe(false);
    expect(host.dataset.fcMacAutomationResult).toBeUndefined();
  });

  it('rolls back a native add failure during hyperlink undo', async () => {
    const native = await makeNativeInstance(host);
    const { instance, workbook, store, history, controller } = native;
    expect(
      workbook.addHyperlinkRange(0, 74, 1, 74, 1, 'https://undo.example', 'Undo', '', ''),
    ).toBe(true);
    const before = workbook.getHyperlinksFull(0);
    controller.setPolicy({ defaultOperation: 'deny', operations: { hyperlink: true } });

    await runMacAutomationScript(instance, 'removeHyperlinks');
    expect(workbook.getHyperlinksFull(0)).toEqual([]);
    expect(history.canUndo()).toBe(true);
    const resultBefore = host.dataset.fcMacAutomationResult;
    const add = workbook.addHyperlinkRange.bind(workbook);
    workbook.addHyperlinkRange = () => false;

    let caught: unknown;
    try {
      history.undo();
    } catch (error) {
      caught = error;
    }

    expect(caught).toBeInstanceOf(Error);
    expect((caught as Error).message).toContain('Could not restore hyperlink');
    expect(workbook.getHyperlinksFull(0)).toEqual([]);
    expect(store.getState().format.formats.get('0:74:1')?.hyperlink).toBeUndefined();
    expect(history.canUndo()).toBe(true);
    expect(host.dataset.fcMacAutomationResult).toBe(resultBefore);

    workbook.addHyperlinkRange = add;
    expect(history.undo()).toBe(true);
    expect(workbook.getHyperlinksFull(0)).toEqual(before);
    expect(store.getState().format.formats.get('0:74:1')?.hyperlink).toBeUndefined();
  });

  it('disposes open gallery and result overlays with the instance', async () => {
    const instance = makeInstance(host);
    const result = runMacAutomationScript(instance, 'countEmptyRows');
    expect(document.querySelector('.fc-macautomationresult')).not.toBeNull();
    disposeMacAutomationDialogs(instance);
    await result;
    expect(document.querySelector('.fc-macautomationresult')).toBeNull();

    const gallery = openMacAutomationGallery(instance);
    expect(document.querySelector('.fc-macautomationdlg')).not.toBeNull();
    disposeMacAutomationDialogs(instance);
    await gallery;
    expect(document.querySelector('.fc-macautomationdlg')).toBeNull();
  });
});
