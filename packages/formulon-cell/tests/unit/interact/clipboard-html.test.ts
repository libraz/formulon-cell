import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachClipboard, type ClipboardHandle } from '../../../src/interact/clipboard.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const EXTERNAL_HTML = '<table><tr><td><b>bold</b></td><td style="color:red">red</td></tr></table>';

const firePaste = (host: HTMLElement, data: Record<string, string>): ClipboardEvent => {
  const transfer = new DataTransfer();
  for (const [type, value] of Object.entries(data)) transfer.setData(type, value);
  const event = new ClipboardEvent('paste', {
    clipboardData: transfer as unknown as DataTransfer,
    bubbles: true,
    cancelable: true,
  });
  host.dispatchEvent(event);
  return event;
};

describe('clipboard HTML and implicit formats', () => {
  let wb: WorkbookHandle;
  let store: SpreadsheetStore;
  let host: HTMLElement;
  let handle: ClipboardHandle;
  const fmt = (row: number, col: number) =>
    store.getState().format.formats.get(addrKey({ sheet: 0, row, col }));

  beforeEach(async () => {
    wb = await WorkbookHandle.createDefault({ preferStub: true });
    store = createSpreadsheetStore();
    host = document.createElement('div');
    document.body.appendChild(host);
    handle = attachClipboard({ host, store, wb, onAfterCommit: () => {} });
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
  });

  afterEach(() => {
    handle.detach();
    host.remove();
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it('pastes an external HTML table with its bold before falling back to plain text', () => {
    firePaste(host, { 'text/plain': 'bold\tred', 'text/html': EXTERNAL_HTML });
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'text', value: 'red' });
    expect(fmt(0, 0)?.bold).toBe(true);
  });

  it('pastes our own HTML through the lossless plain text instead', () => {
    // Our HTML carries displayed numbers; the plain text keeps full precision.
    wb.setNumber({ sheet: 0, row: 5, col: 0 }, 1.23456);
    mutators.replaceCells(store, wb.cells(0));
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 5, col: 0 },
      { numFmt: { kind: 'fixed', decimals: 2 } },
    );
    mutators.setRange(store, { sheet: 0, r0: 5, c0: 0, r1: 5, c1: 0 });
    const transfer = new DataTransfer();
    host.dispatchEvent(
      new ClipboardEvent('copy', { clipboardData: transfer, bubbles: true, cancelable: true }),
    );
    const html = transfer.getData('text/html');
    expect(html).toContain('1.23<');
    // A later paste, after the internal copy is gone.
    mutators.setCopyRange(store, null);
    mutators.setActive(store, { sheet: 0, row: 0, col: 2 });
    firePaste(host, { 'text/plain': transfer.getData('text/plain'), 'text/html': html });
    expect(wb.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({ kind: 'number', value: 1.23456 });
  });

  it('gives a pasted General cell the format its text implies', () => {
    firePaste(host, { 'text/plain': '1,234' });
    expect(fmt(0, 0)?.numFmt).toEqual({ kind: 'fixed', decimals: 0, thousands: true });
  });

  it('writes text/plain and text/html from the copy shortcut', async () => {
    const write = vi.fn().mockResolvedValue(undefined);
    const writeText = vi.fn().mockResolvedValue(undefined);
    class FakeClipboardItem {
      constructor(readonly items: Record<string, Blob>) {}
    }
    vi.stubGlobal('ClipboardItem', FakeClipboardItem);
    vi.stubGlobal('navigator', { clipboard: { write, writeText } });
    wb.setText({ sheet: 0, row: 0, col: 0 }, 'x');
    mutators.replaceCells(store, wb.cells(0));
    await handle.runShortcut('copy');
    expect(writeText).not.toHaveBeenCalled();
    const item = write.mock.calls[0]?.[0]?.[0] as FakeClipboardItem;
    expect(Object.keys(item.items).sort()).toEqual(['text/html', 'text/plain']);
    expect(await item.items['text/plain']?.text()).toBe('x');
    expect(await item.items['text/html']?.text()).toContain('<td>x</td>');
  });

  it('falls back to writeText without ClipboardItem', async () => {
    const writeText = vi.fn().mockResolvedValue(undefined);
    vi.stubGlobal('ClipboardItem', undefined);
    vi.stubGlobal('navigator', { clipboard: { writeText } });
    wb.setText({ sheet: 0, row: 0, col: 0 }, 'x');
    mutators.replaceCells(store, wb.cells(0));
    await handle.runShortcut('copy');
    expect(writeText).toHaveBeenCalledWith('x');
  });

  it('pastes HTML read by the paste shortcut', async () => {
    const blob = (type: string, text: string) => new Blob([text], { type });
    const read = vi.fn().mockResolvedValue([
      {
        types: ['text/plain', 'text/html'],
        getType: (type: string) =>
          Promise.resolve(
            type === 'text/html' ? blob(type, EXTERNAL_HTML) : blob(type, 'bold\tred'),
          ),
      },
    ]);
    vi.stubGlobal('navigator', { clipboard: { read, readText: vi.fn() } });
    await handle.runShortcut('paste');
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'bold' });
    expect(fmt(0, 0)?.bold).toBe(true);
  });
});
