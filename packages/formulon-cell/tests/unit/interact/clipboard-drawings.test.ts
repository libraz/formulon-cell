import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { createSessionChart, sessionChartById } from '../../../src/commands/session-chart.js';
import {
  createSessionShape,
  sessionIllustrationById,
} from '../../../src/commands/session-illustration.js';
import { cellOrigin, sessionPosition } from '../../../src/engine/drawing-geometry.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachClipboard, type ClipboardHandle } from '../../../src/interact/clipboard.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const range = (r0: number, c0: number, r1: number, c1: number) => ({ sheet: 0, r0, c0, r1, c1 });

/** Leave any selected object for cell (row, col) on `sheet`. */
const select = (store: SpreadsheetStore, sheet: number, row: number, col: number): void => {
  mutators.setSelectedObject(store, null);
  store.setState((s) => ({
    ...s,
    selection: {
      ...s.selection,
      active: { sheet, row, col },
      anchor: { sheet, row, col },
      range: { sheet, r0: row, c0: col, r1: row, c1: col },
    },
  }));
};

describe('clipboard with a selected chart or picture', () => {
  let wb: WorkbookHandle;
  let store: SpreadsheetStore;
  let history: History;
  let host: HTMLElement;
  let handle: ClipboardHandle;
  let systemText: string;
  let readText: ReturnType<typeof vi.fn>;

  beforeEach(async () => {
    wb = await WorkbookHandle.createDefault();
    wb.setText({ sheet: 0, row: 0, col: 1 }, 'North');
    wb.setText({ sheet: 0, row: 1, col: 0 }, 'Jan');
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 10);
    wb.setText({ sheet: 0, row: 2, col: 0 }, 'Feb');
    wb.setNumber({ sheet: 0, row: 2, col: 1 }, 15);
    store = createSpreadsheetStore();
    mutators.replaceCells(store, wb.cells(0));
    history = new History();
    host = document.createElement('div');
    document.body.appendChild(host);
    systemText = 'Earlier text';
    readText = vi.fn(async () => systemText);
    vi.stubGlobal('navigator', {
      clipboard: {
        readText,
        writeText: vi.fn(async (text: string) => {
          systemText = text;
        }),
      },
    });
    handle = attachClipboard({ host, store, wb, history, onAfterCommit: () => {} });
  });

  afterEach(() => {
    handle.detach();
    vi.unstubAllGlobals();
    host.remove();
    wb.dispose();
  });

  const chart = () => {
    const created = createSessionChart(store, range(0, 0, 2, 1), {
      wb,
      kind: 'column',
      anchor: { row: 1, col: 4 },
    });
    if (!created) throw new Error('chart was not created');
    mutators.setSelectedObject(store, { id: created.id, name: 'Chart 1' });
    return created;
  };

  const charts = () => store.getState().charts.charts;

  it('pastes a copied chart at the active cell and selects it', async () => {
    const original = chart();

    await handle.runShortcut('copy');
    expect(systemText).toBe('Chart 1');
    expect(charts()).toHaveLength(1);

    select(store, 0, 6, 2);
    await handle.runShortcut('paste');
    const [, pasted] = charts();
    if (!pasted) throw new Error('chart was not pasted');
    expect(pasted.nativeObjectId).not.toBe(original.nativeObjectId);
    expect(pasted.data).toEqual(original.data);
    const at = sessionPosition(store, pasted.nativeAnchor);
    const cell = cellOrigin(store, 6, 2);
    expect(at.x).toBeCloseTo(cell.x, 0);
    expect(at.y).toBeCloseTo(cell.y, 0);
    expect(at.w).toBeCloseTo(sessionPosition(store, original.nativeAnchor).w, 0);
    expect(store.getState().ui.selectedObject?.id).toBe(pasted.id);
    // The cells under the chart are untouched.
    expect(wb.getValue({ sheet: 0, row: 6, col: 2 }).kind).toBe('blank');

    expect(history.undo()).toBe(true);
    expect(charts().map((entry) => entry.id)).toEqual([original.id]);
  });

  it('pastes a chart on another sheet, still charting the sheet it came from', async () => {
    const original = chart();
    const second = wb.addSheet('Second');
    await handle.runShortcut('copy');

    select(store, second, 0, 0);
    await handle.runShortcut('paste');
    const pasted = charts().find((entry) => entry.sheet === second);
    if (!pasted) throw new Error('chart was not pasted');
    expect(pasted.source).toEqual(original.source);
    expect(pasted.data).toEqual(original.data);
    expect(wb.listDrawingObjects(second)).toHaveLength(1);
  });

  it('cuts a chart off the sheet at once and pastes it back where asked', async () => {
    const original = chart();

    await handle.runShortcut('cut');
    expect(charts()).toEqual([]);
    expect(wb.listDrawingObjects(0)).toEqual([]);

    select(store, 0, 3, 0);
    await handle.runShortcut('paste');
    const [pasted] = charts();
    if (!pasted) throw new Error('chart was not pasted');
    expect(pasted.data).toEqual(original.data);
    expect(sessionPosition(store, pasted.nativeAnchor).y).toBeCloseTo(cellOrigin(store, 3, 0).y, 0);

    expect(history.undo()).toBe(true);
    expect(charts()).toEqual([]);
    expect(history.undo()).toBe(true);
    expect(charts().map((entry) => entry.id)).toEqual([original.id]);
  });

  it('pastes text copied elsewhere rather than the chart copied before it', async () => {
    chart();
    await handle.runShortcut('copy');
    systemText = 'Outside';

    select(store, 0, 6, 0);
    await handle.runShortcut('paste');
    expect(charts()).toHaveLength(1);
    expect(wb.getValue({ sheet: 0, row: 6, col: 0 })).toEqual({ kind: 'text', value: 'Outside' });
  });

  it('forgets the chart once cells are copied', async () => {
    chart();
    await handle.runShortcut('copy');
    select(store, 0, 1, 1);
    await handle.runShortcut('copy');
    // A browser that withholds the clipboard leaves only the local copy to paste.
    readText.mockRejectedValue(new DOMException('Denied', 'NotAllowedError'));

    select(store, 0, 6, 1);
    await handle.runShortcut('paste');
    expect(charts()).toHaveLength(1);
    expect(wb.getValue({ sheet: 0, row: 6, col: 1 })).toEqual({ kind: 'number', value: 10 });
  });

  it('pastes the copied chart when the browser withholds the clipboard', async () => {
    chart();
    await handle.runShortcut('copy');
    readText.mockRejectedValue(new DOMException('Denied', 'NotAllowedError'));

    select(store, 0, 6, 0);
    await handle.runShortcut('paste');
    expect(charts()).toHaveLength(2);
  });

  it('pastes nothing onto a protected sheet and cuts nothing from one', async () => {
    const original = chart();
    await handle.runShortcut('copy');
    setProtectedSheet(store, 0, true);

    select(store, 0, 6, 0);
    await handle.runShortcut('paste');
    expect(charts()).toHaveLength(1);

    mutators.setSelectedObject(store, { id: original.id, name: 'Chart 1' });
    await handle.runShortcut('cut');
    expect(sessionChartById(store.getState(), original.id)).not.toBeNull();
  });

  it('copies a shape through the copy and paste events', () => {
    createSessionShape(store, range(0, 0, 0, 0), {
      id: 'shape-1',
      shape: 'rectangle',
      x: 20,
      y: 30,
      w: 160,
      h: 96,
    });
    mutators.setSelectedObject(store, { id: 'shape-1', name: 'Rectangle 1' });
    const transfer = new DataTransfer();
    const copyEvent = new ClipboardEvent('copy', {
      clipboardData: transfer,
      bubbles: true,
      cancelable: true,
    });
    host.dispatchEvent(copyEvent);
    expect(copyEvent.defaultPrevented).toBe(true);
    expect(transfer.getData('text/plain')).toBe('Rectangle 1');

    select(store, 0, 4, 2);
    const pasteEvent = new ClipboardEvent('paste', {
      clipboardData: transfer,
      bubbles: true,
      cancelable: true,
    });
    host.dispatchEvent(pasteEvent);
    expect(pasteEvent.defaultPrevented).toBe(true);
    const pasted = sessionIllustrationById(store.getState(), 'shape-1-copy');
    expect(pasted).toMatchObject({ shape: 'rectangle', w: 160, h: 96, ...cellOrigin(store, 4, 2) });
    expect(store.getState().ui.selectedObject?.id).toBe('shape-1-copy');
    expect(history.undo()).toBe(true);
    expect(sessionIllustrationById(store.getState(), 'shape-1-copy')).toBeNull();
  });
});
