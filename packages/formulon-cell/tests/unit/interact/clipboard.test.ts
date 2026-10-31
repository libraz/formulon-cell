import { afterEach, beforeEach, describe, expect, it, type Mock, vi } from 'vitest';
import { pasteTSV } from '../../../src/commands/clipboard/paste.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { fixedFormPolicy } from '../../../src/commands/interaction-policy.js';
import { insertRows } from '../../../src/commands/structure.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachClipboard } from '../../../src/interact/clipboard.js';
import type { PasteOptionsActivation } from '../../../src/interact/paste-options.js';
import {
  type CellFormat,
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const seedAndMirror = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  cells: Array<{ row: number; col: number; value: number | string | boolean; formula?: string }>,
): void => {
  store.setState((s) => {
    const map = new Map(s.data.cells);
    for (const c of cells) {
      const addr = { sheet: 0, row: c.row, col: c.col };
      if (c.formula) {
        wb.setFormula(addr, c.formula);
        map.set(addrKey(addr), {
          value:
            typeof c.value === 'number'
              ? { kind: 'number', value: c.value }
              : typeof c.value === 'boolean'
                ? { kind: 'bool', value: c.value }
                : { kind: 'text', value: c.value },
          formula: c.formula,
        });
      } else if (typeof c.value === 'number') {
        wb.setNumber(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'number', value: c.value }, formula: null });
      } else if (typeof c.value === 'boolean') {
        wb.setBool(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'bool', value: c.value }, formula: null });
      } else {
        wb.setText(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'text', value: c.value }, formula: null });
      }
    }
    return { ...s, data: { ...s.data, cells: map } };
  });
  wb.recalc();
};

const setRange = (
  store: SpreadsheetStore,
  r0: number,
  c0: number,
  r1: number,
  c1: number,
  sheet = 0,
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      ...s.selection,
      active: { sheet, row: r0, col: c0 },
      anchor: { sheet, row: r0, col: c0 },
      range: { sheet, r0, c0, r1, c1 },
    },
  }));
};

const setFormat = (store: SpreadsheetStore, row: number, col: number, format: CellFormat): void => {
  store.setState((s) => {
    const formats = new Map(s.format.formats);
    formats.set(addrKey({ sheet: 0, row, col }), format);
    return { ...s, format: { ...s.format, formats } };
  });
};

const formatAt = (store: SpreadsheetStore, row: number, col: number): CellFormat | undefined =>
  store.getState().format.formats.get(addrKey({ sheet: 0, row, col }));

const fireClipboard = (
  host: HTMLElement,
  type: 'copy' | 'cut' | 'paste',
  data?: string,
): { event: ClipboardEvent; transfer: DataTransfer } => {
  const transfer = new DataTransfer();
  if (data !== undefined) transfer.setData('text/plain', data);
  const event = new ClipboardEvent(type, {
    clipboardData: transfer as unknown as DataTransfer,
    bubbles: true,
    cancelable: true,
  });
  host.dispatchEvent(event);
  return { event, transfer };
};

describe('attachClipboard', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let onAfterCommit: Mock<() => void>;
  let unregisterController: (() => void) | null;

  beforeEach(async () => {
    host = document.createElement('div');
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
    onAfterCommit = vi.fn<() => void>();
    unregisterController = null;
  });

  afterEach(() => {
    unregisterController?.();
    document.body.innerHTML = '';
  });

  it('copy writes TSV to the clipboard and captures a snapshot', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 'two' },
    ]);
    setRange(store, 0, 0, 0, 1);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    const { event, transfer } = fireClipboard(host, 'copy');
    expect(transfer.getData('text/plain')).toBe('1\ttwo');
    expect(event.defaultPrevented).toBe(true);
    expect(handle.getSnapshot()).not.toBeNull();
    expect(handle.getSnapshot()?.rows).toBe(1);
    expect(handle.getSnapshot()?.cols).toBe(2);
    expect(store.getState().ui.copyRange).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    expect(onAfterCommit).not.toHaveBeenCalled();
    handle.detach();
  });

  it.each([
    ['windows-de_DE', '1,5\tWAHR\tFALSCH'],
    ['windows-fr_FR', '1,5\tVRAI\tFAUX'],
  ] as const)('copies edit-input text for the live %s profile', async (profile, expectedTsv) => {
    const localizedHost = document.createElement('div');
    document.body.appendChild(localizedHost);
    const localizedStore = createSpreadsheetStore();
    const localizedWb = await WorkbookHandle.createDefault();
    const localizedAfterCommit = vi.fn<() => void>();
    try {
      expect(localizedWb.setSpreadsheetProfileId(profile)).toBe(true);
      seedAndMirror(localizedStore, localizedWb, [
        { row: 0, col: 0, value: 1.5 },
        { row: 0, col: 1, value: true },
        { row: 0, col: 2, value: false },
      ]);
      setRange(localizedStore, 0, 0, 0, 2);
      const localizedHandle = attachClipboard({
        host: localizedHost,
        store: localizedStore,
        wb: localizedWb,
        onAfterCommit: localizedAfterCommit,
      });

      const { transfer } = fireClipboard(localizedHost, 'copy');
      expect(transfer.getData('text/plain')).toBe(expectedTsv);
      expect(transfer.getData('text/html')).toContain(expectedTsv.split('\t')[0]);
      expect(transfer.getData('text/html')).toContain(expectedTsv.split('\t')[1] ?? '');

      setRange(localizedStore, 1, 0, 1, 2);
      pasteTSV(localizedStore.getState(), localizedWb, transfer.getData('text/plain'));
      expect(localizedWb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
        kind: 'number',
        value: 1.5,
      });
      expect(localizedWb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({
        kind: 'bool',
        value: true,
      });
      expect(localizedWb.getValue({ sheet: 0, row: 1, col: 2 })).toEqual({
        kind: 'bool',
        value: false,
      });
      localizedHandle.detach();
    } finally {
      localizedWb.dispose();
      localizedHost.remove();
    }
  });

  it('copy is a no-op while the editor is active', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 1 }]);
    setRange(store, 0, 0, 0, 0);
    store.setState((s) => ({
      ...s,
      ui: { ...s.ui, editor: { kind: 'edit', raw: 'x', caret: 0 } },
    }));
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    const { event, transfer } = fireClipboard(host, 'copy');
    expect(transfer.getData('text/plain')).toBe('');
    expect(event.defaultPrevented).toBe(false);
    expect(handle.getSnapshot()).toBeNull();
    handle.detach();
  });

  it('copy clears a stale copy marquee when the next ordinary range is too large', () => {
    mutators.setCopyRanges(store, [
      { sheet: 0, r0: 2, c0: 0, r1: 2, c1: 16383 },
      { sheet: 0, r0: 4, c0: 0, r1: 4, c1: 16383 },
    ]);
    setRange(store, 0, 0, 1_000, 999);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    fireClipboard(host, 'copy');

    expect(store.getState().ui.copyRange).toBeNull();
    expect(store.getState().ui.copyRanges).toBeNull();
    handle.detach();
  });

  it('replaces row copy marquees when copying a column next', () => {
    seedAndMirror(store, wb, [
      { row: 2, col: 1, value: 'row' },
      { row: 0, col: 3, value: 'col' },
    ]);
    setRange(store, 2, 0, 2, 16383);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    fireClipboard(host, 'copy');
    expect(store.getState().ui.copyRange).toEqual({ sheet: 0, r0: 2, c0: 0, r1: 2, c1: 16383 });

    setRange(store, 0, 3, 1048575, 3);
    fireClipboard(host, 'copy');

    expect(store.getState().ui.copyRange).toEqual({ sheet: 0, r0: 0, c0: 3, r1: 1048575, c1: 3 });
    expect(store.getState().ui.copyRanges).toBeNull();
    handle.detach();
  });

  it('copy uses the trimmed payload for whole-column HTML and structured snapshots', () => {
    seedAndMirror(store, wb, [
      { row: 4, col: 2, value: 'top' },
      { row: 6, col: 2, value: 'bottom' },
    ]);
    setRange(store, 0, 2, 1048575, 2);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    const { transfer } = fireClipboard(host, 'copy');

    expect(transfer.getData('text/plain')).toBe('top\r\n\r\nbottom');
    expect(transfer.getData('text/html')).toContain('<table>');
    expect(transfer.getData('text/html')).toContain('top');
    expect(transfer.getData('text/html')).toContain('bottom');
    expect(handle.getSnapshot()?.range).toEqual({ sheet: 0, r0: 4, c0: 2, r1: 6, c1: 2 });
    expect(handle.getSnapshot()?.rows).toBe(3);
    expect(handle.getSnapshot()?.cols).toBe(1);
    expect(store.getState().ui.copyRange).toEqual({ sheet: 0, r0: 0, c0: 2, r1: 1048575, c1: 2 });
    handle.detach();
  });

  it('copy keeps format-only cells in trimmed whole-column HTML and snapshots', () => {
    setFormat(store, 9, 2, { hyperlink: 'https://example.test', bold: true });
    setRange(store, 0, 2, 1048575, 2);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    const { transfer } = fireClipboard(host, 'copy');

    expect(transfer.getData('text/plain')).toBe('');
    expect(transfer.getData('text/html')).toContain('href="https://example.test"');
    expect(transfer.getData('text/html')).toContain('font-weight:bold');
    expect(handle.getSnapshot()?.range).toEqual({ sheet: 0, r0: 9, c0: 2, r1: 9, c1: 2 });
    expect(handle.getSnapshot()?.cells[0]?.[0]?.format).toMatchObject({
      hyperlink: 'https://example.test',
      bold: true,
    });
    handle.detach();
  });

  it('refreshes a live whole-column copy from its full range without scanning the sheet', () => {
    seedAndMirror(store, wb, [{ row: 4, col: 2, value: 'top' }]);
    setRange(store, 0, 2, 1_048_575, 2);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });
    fireClipboard(host, 'copy');

    wb.setText({ sheet: 0, row: 20_000, col: 2 }, 'added after copy');
    const ranges: Array<{ sheet: number; r0: number; c0: number; r1: number; c1: number }> = [];
    vi.spyOn(wb, 'cellsInRange').mockImplementation(function* (range) {
      ranges.push({ ...range });
      yield {
        addr: { sheet: 0, row: 4, col: 2 },
        value: { kind: 'text', value: 'top' },
        formula: null,
      };
      yield {
        addr: { sheet: 0, row: 20_000, col: 2 },
        value: { kind: 'text', value: 'added after copy' },
        formula: null,
      };
    });
    const fullSheetScans = vi.spyOn(wb, 'cells').mockImplementation(() => {
      throw new Error('unexpected full-sheet source scan');
    });

    const refreshed = handle.getSnapshot();

    expect(ranges).toEqual([{ sheet: 0, r0: 0, c0: 2, r1: 1_048_575, c1: 2 }]);
    expect(fullSheetScans).not.toHaveBeenCalled();
    expect(refreshed?.range).toEqual({ sheet: 0, r0: 4, c0: 2, r1: 20_000, c1: 2 });
    expect(refreshed?.cells[19_996]?.[0]?.value).toEqual({
      kind: 'text',
      value: 'added after copy',
    });
    handle.detach();
  });

  it('refreshes copied formulas without generating unused localized TSV', () => {
    seedAndMirror(store, wb, [{ row: 4, col: 2, value: 3, formula: '=SUM(1,2)' }]);
    setRange(store, 0, 2, 1_048_575, 2);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });
    fireClipboard(host, 'copy');

    wb.setFormula({ sheet: 0, row: 4, col: 2 }, '=SUM(3,4)');
    wb.setText({ sheet: 0, row: 20_000, col: 2 }, 'added after copy');
    const localize = vi.spyOn(wb, 'localizeFormulaText');

    const refreshed = handle.getSnapshot();

    expect(localize).not.toHaveBeenCalled();
    expect(refreshed?.range).toEqual({ sheet: 0, r0: 4, c0: 2, r1: 20_000, c1: 2 });
    expect(refreshed?.cells[0]?.[0]?.formula).toBe('=SUM(3,4)');
    expect(refreshed?.cells[19_996]?.[0]?.value).toEqual({
      kind: 'text',
      value: 'added after copy',
    });
    handle.detach();
  });

  it('captures internal paste undo snapshots without localizing destination formulas', () => {
    const history = new History();
    const onPasteOptions = vi.fn<(activation: PasteOptionsActivation) => void>();
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 3, formula: '=SUM(1,2)' },
      { row: 1, col: 0, value: 5, formula: '=SUM(2,3)' },
    ]);
    setFormat(store, 0, 0, { italic: true, fill: '#e2f0d9' });
    setFormat(store, 1, 0, { bold: true, fill: '#fff2cc' });
    setRange(store, 0, 0, 0, 0);
    const handle = attachClipboard({ host, history, store, wb, onAfterCommit, onPasteOptions });
    const copied = fireClipboard(host, 'copy').transfer.getData('text/plain');
    const localize = vi.spyOn(wb, 'localizeFormulaText');
    setRange(store, 1, 0, 1, 0);

    fireClipboard(host, 'paste', copied);

    expect(localize).not.toHaveBeenCalledWith('=SUM(2,3)');
    expect(onPasteOptions).toHaveBeenCalledTimes(1);
    const before = onPasteOptions.mock.calls[0]?.[0].before;
    expect(before?.cells[0]?.[0]?.formula).toBe('=SUM(2,3)');
    expect(before?.cells[0]?.[0]?.format).toMatchObject({ bold: true, fill: '#fff2cc' });
    expect(wb.cellFormula({ sheet: 0, row: 1, col: 0 })).toBe('=SUM(1,2)');
    expect(formatAt(store, 1, 0)).toMatchObject({ italic: true, fill: '#e2f0d9' });
    expect(onAfterCommit).toHaveBeenCalled();
    expect(history.undo()).toBe(true);
    wb.recalc();
    expect(wb.cellFormula({ sheet: 0, row: 1, col: 0 })).toBe('=SUM(2,3)');
    expect(formatAt(store, 1, 0)).toMatchObject({ bold: true, fill: '#fff2cc' });
    handle.detach();
  });

  it('shares format reads across copy snapshot and HTML, then refreshes on the next copy', async () => {
    const formatHost = document.createElement('div');
    document.body.appendChild(formatHost);
    const formatStore = createSpreadsheetStore();
    const formatWb = await WorkbookHandle.createDefault();
    let formatHandle: ReturnType<typeof attachClipboard> | null = null;
    try {
      expect(formatWb.capabilities.effectiveStyle).toBe(true);
      const defaultXf = formatWb.getCellStyleXf(0);
      if (!defaultXf) throw new Error('missing default XF');
      const defaultFont = formatWb.getFontRecord(defaultXf.fontIndex);
      if (!defaultFont) throw new Error('missing default font');
      const boldFont = formatWb.addFontRecord({ ...defaultFont, bold: true, hasBold: true });
      const boldXf = formatWb.addXfRecord({ ...defaultXf, fontIndex: boldFont });
      const italicFont = formatWb.addFontRecord({ ...defaultFont, italic: true, hasItalic: true });
      const italicXf = formatWb.addXfRecord({ ...defaultXf, fontIndex: italicFont });
      expect(formatWb.setColumnStyle(0, 0, 0, boldXf)).toBe(true);
      expect(formatWb.setCellXfIndex(0, 0, 0, italicXf)).toBe(true);
      for (let row = 0; row < 1000; row += 1) formatWb.setNumber({ sheet: 0, row, col: 0 }, row);
      mutators.replaceCells(formatStore, formatWb.cells(0));
      formatStore.setState((state) => ({
        ...state,
        format: {
          ...state.format,
          formats: new Map([[addrKey({ sheet: 0, row: 1, col: 0 }), { italic: true }]]),
        },
      }));
      setRange(formatStore, 0, 0, 999, 0);

      const effectiveReads = vi.spyOn(formatWb, 'getEffectiveStyle');
      const xfReads = vi.spyOn(formatWb, 'getCellXf');
      const fontReads = vi.spyOn(formatWb, 'getResolvedFontRecord');
      const fillReads = vi.spyOn(formatWb, 'getResolvedFillRecord');
      const borderReads = vi.spyOn(formatWb, 'getResolvedBorderRecord');
      formatHandle = attachClipboard({
        host: formatHost,
        store: formatStore,
        wb: formatWb,
        onAfterCommit,
      });

      const first = fireClipboard(formatHost, 'copy');
      expect(first.transfer.getData('text/html')).toContain('font-weight:bold');
      expect(first.transfer.getData('text/html')).toContain('font-style:italic');
      expect(effectiveReads).toHaveBeenCalledTimes(1000);
      expect(xfReads).toHaveBeenCalledTimes(2);
      // Each unique XF reads its font plus the workbook display font.
      expect(fontReads).toHaveBeenCalledTimes(4);
      expect(fillReads).toHaveBeenCalledTimes(2);
      expect(borderReads).toHaveBeenCalledTimes(2);

      const firstSnapshot = formatHandle.getSnapshot();
      expect(firstSnapshot?.cells[0]?.[0]?.format).toMatchObject({ italic: true });
      expect(firstSnapshot?.cells[1]?.[0]?.format).toMatchObject({ bold: true, italic: true });
      expect(firstSnapshot?.cells[2]?.[0]?.format).toMatchObject({ bold: true });

      for (let row = 0; row < 1000; row += 1) {
        expect(formatWb.setCellXfIndex(0, row, 0, italicXf)).toBe(true);
      }
      effectiveReads.mockClear();
      xfReads.mockClear();
      fontReads.mockClear();
      fillReads.mockClear();
      borderReads.mockClear();
      const second = fireClipboard(formatHost, 'copy');
      const secondHtml = second.transfer.getData('text/html');
      expect(secondHtml).toContain('font-style:italic');
      expect(secondHtml).not.toContain('font-weight:bold');
      expect(effectiveReads).toHaveBeenCalledTimes(1000);
      expect(xfReads).toHaveBeenCalledTimes(1);
      expect(fontReads).toHaveBeenCalledTimes(2);
      expect(fillReads).toHaveBeenCalledTimes(1);
      expect(borderReads).toHaveBeenCalledTimes(1);

      const secondSnapshot = formatHandle.getSnapshot();
      expect(secondSnapshot?.cells[2]?.[0]?.format).toMatchObject({ italic: true });
      expect(secondSnapshot?.cells[2]?.[0]?.format?.bold).toBeUndefined();
    } finally {
      formatHandle?.detach();
      formatWb.dispose();
      formatHost.remove();
    }
  });

  it('pastes the internal snapshot for format-only copies with empty plain text', () => {
    setFormat(store, 9, 2, { hyperlink: 'https://example.test', bold: true });
    setRange(store, 0, 2, 1048575, 2);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    const { transfer } = fireClipboard(host, 'copy');
    setRange(store, 0, 4, 0, 4);
    const { event } = fireClipboard(host, 'paste', transfer.getData('text/plain'));

    expect(event.defaultPrevented).toBe(true);
    expect(formatAt(store, 9, 4)).toMatchObject({
      hyperlink: 'https://example.test',
      bold: true,
    });
    expect(store.getState().selection.range).toEqual({
      sheet: 0,
      r0: 0,
      c0: 4,
      r1: 1_048_575,
      c1: 4,
    });
    expect(onAfterCommit).toHaveBeenCalledTimes(1);
    handle.detach();
  });

  it('cut writes TSV and snapshots while leaving the source intact', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 5 },
      { row: 0, col: 1, value: 6 },
    ]);
    setFormat(store, 0, 0, { bold: true, fill: '#fff2cc' });
    setFormat(store, 0, 1, { italic: true });
    setRange(store, 0, 0, 0, 1);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    const { event, transfer } = fireClipboard(host, 'cut');
    expect(transfer.getData('text/plain')).toBe('5\t6');
    expect(event.defaultPrevented).toBe(true);
    expect(handle.getSnapshot()).not.toBeNull();
    expect(store.getState().ui.copyRange).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 5 });
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 6 });
    expect(formatAt(store, 0, 0)).toMatchObject({ bold: true, fill: '#fff2cc' });
    expect(formatAt(store, 0, 1)).toMatchObject({ italic: true });
    expect(onAfterCommit).not.toHaveBeenCalled();
    handle.detach();
  });

  it('cut uses the trimmed payload for whole-column structured snapshots', () => {
    seedAndMirror(store, wb, [
      { row: 4, col: 2, value: 'top' },
      { row: 6, col: 2, value: 'bottom' },
    ]);
    setFormat(store, 4, 2, { bold: true });
    setFormat(store, 6, 2, { italic: true });
    setRange(store, 0, 2, 1048575, 2);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    const { transfer } = fireClipboard(host, 'cut');

    expect(transfer.getData('text/plain')).toBe('top\r\n\r\nbottom');
    expect(handle.getSnapshot()?.range).toEqual({ sheet: 0, r0: 4, c0: 2, r1: 6, c1: 2 });
    expect(handle.getSnapshot()?.rows).toBe(3);
    expect(handle.getSnapshot()?.cols).toBe(1);
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 4, col: 2 })).toEqual({ kind: 'text', value: 'top' });
    expect(wb.getValue({ sheet: 0, row: 6, col: 2 })).toEqual({ kind: 'text', value: 'bottom' });
    expect(formatAt(store, 4, 2)).toMatchObject({ bold: true });
    expect(formatAt(store, 6, 2)).toMatchObject({ italic: true });
    expect(store.getState().ui.copyRange).toEqual({ sheet: 0, r0: 0, c0: 2, r1: 1048575, c1: 2 });
    handle.detach();
  });

  it('rejects a whole-sheet cut without disturbing the previous copy payload', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 7 }]);
    setFormat(store, 0, 0, { bold: true, fill: '#fff2cc' });
    setRange(store, 0, 0, 0, 0);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });
    const { transfer } = fireClipboard(host, 'copy');
    const marquee = store.getState().ui.copyRange;

    setRange(store, 0, 0, 1048575, 16383);
    const { event } = fireClipboard(host, 'cut');

    expect(event.defaultPrevented).toBe(true);
    expect(store.getState().ui.copyRange).toEqual(marquee);
    expect(store.getState().ui.copyMode).toBe('copy');
    expect(handle.getSnapshot()).not.toBeNull();
    setRange(store, 2, 2, 2, 2);
    fireClipboard(host, 'paste', transfer.getData('text/plain'));
    expect(formatAt(store, 2, 2)).toMatchObject({ bold: true, fill: '#fff2cc' });
    handle.detach();
  });

  it('rejects a whole-sheet cut shortcut without disturbing the previous copy payload', async () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 7 }]);
    setFormat(store, 0, 0, { bold: true, fill: '#fff2cc' });
    setRange(store, 0, 0, 0, 0);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });
    await handle.runShortcut('copy');
    const marquee = store.getState().ui.copyRange;

    setRange(store, 0, 0, 1048575, 16383);
    await handle.runShortcut('cut');

    expect(store.getState().ui.copyRange).toEqual(marquee);
    expect(handle.getSnapshot()).not.toBeNull();
    setRange(store, 2, 2, 2, 2);
    await handle.runShortcut('paste');
    expect(formatAt(store, 2, 2)).toMatchObject({ bold: true, fill: '#fff2cc' });
    handle.detach();
  });

  it('does not mutate data or history until a cut is pasted', () => {
    const history = new History();
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 5 }]);
    setFormat(store, 0, 0, { bold: true, fill: '#fff2cc' });
    setRange(store, 0, 0, 0, 0);
    const handle = attachClipboard({ host, history, store, wb, onAfterCommit });

    fireClipboard(host, 'cut');
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 5 });
    expect(formatAt(store, 0, 0)).toMatchObject({ bold: true, fill: '#fff2cc' });

    expect(history.undo()).toBe(false);
    handle.detach();
  });

  it('cut is a no-op while the editor is active', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 5 }]);
    setRange(store, 0, 0, 0, 0);
    store.setState((s) => ({
      ...s,
      ui: { ...s.ui, editor: { kind: 'enter', raw: 'x' } },
    }));
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    fireClipboard(host, 'cut');
    expect(onAfterCommit).not.toHaveBeenCalled();
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 5 });
    handle.detach();
  });

  it('paste reads TSV and writes through pasteTSV, calling onAfterCommit', () => {
    setRange(store, 1, 1, 1, 1);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    const { event } = fireClipboard(host, 'paste', 'foo\t42');
    expect(event.defaultPrevented).toBe(true);
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'text', value: 'foo' });
    expect(wb.getValue({ sheet: 0, row: 1, col: 2 })).toEqual({ kind: 'number', value: 42 });
    expect(store.getState().ui.copyRange).toBeNull();
    expect(store.getState().selection.range).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 1, c1: 2 });
    expect(onAfterCommit).toHaveBeenCalledTimes(1);
    handle.detach();
  });

  it('leaves a copy marquee standing after a paste so the source can be reused', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 'a' }]);
    setRange(store, 0, 0, 0, 0);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    const { transfer } = fireClipboard(host, 'copy');
    expect(store.getState().ui.copyMode).toBe('copy');

    setRange(store, 3, 3, 3, 3);
    fireClipboard(host, 'paste', transfer.getData('text/plain'));
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 3, col: 3 })).toEqual({ kind: 'text', value: 'a' });
    expect(store.getState().ui.copyRange).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });

    // Repeat paste into a second destination.
    setRange(store, 5, 5, 5, 5);
    fireClipboard(host, 'paste', transfer.getData('text/plain'));
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 5, col: 5 })).toEqual({ kind: 'text', value: 'a' });
    handle.detach();
  });

  it('repeats a copied 2x2 matrix across a 4x4 selection with one undo step', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 2, formula: '=A1*2' },
      { row: 1, col: 0, value: 3 },
      { row: 1, col: 1, value: 6, formula: '=A2*2' },
    ]);
    setFormat(store, 1, 1, { bold: true, fill: '#fff2cc' });
    setRange(store, 0, 0, 1, 1);
    const history = new History();
    const handle = attachClipboard({ host, history, store, wb, onAfterCommit });

    const { transfer } = fireClipboard(host, 'copy');
    setRange(store, 0, 3, 3, 6);
    fireClipboard(host, 'paste', transfer.getData('text/plain'));
    wb.recalc();

    expect(wb.getValue({ sheet: 0, row: 3, col: 6 })).toEqual({ kind: 'number', value: 6 });
    expect(wb.cellFormula({ sheet: 0, row: 3, col: 6 })).toBe('=F4*2');
    expect(formatAt(store, 3, 6)).toMatchObject({ bold: true, fill: '#fff2cc' });
    expect(store.getState().selection.range).toEqual({ sheet: 0, r0: 0, c0: 3, r1: 3, c1: 6 });

    expect(history.undo()).toBe(true);
    wb.recalc();
    for (let row = 0; row <= 3; row += 1) {
      for (let col = 3; col <= 6; col += 1) {
        expect(wb.getValue({ sheet: 0, row, col }).kind).toBe('blank');
        expect(formatAt(store, row, col)).toBeUndefined();
      }
    }
    expect(history.canUndo()).toBe(false);

    expect(history.redo()).toBe(true);
    wb.recalc();
    expect(wb.cellFormula({ sheet: 0, row: 3, col: 6 })).toBe('=F4*2');
    expect(formatAt(store, 3, 6)).toMatchObject({ bold: true, fill: '#fff2cc' });
    handle.detach();
  });

  it('repeats internal values across a bounded fixed-form selection', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 2 },
      { row: 1, col: 0, value: 3 },
      { row: 1, col: 1, value: 4 },
    ]);
    setRange(store, 0, 0, 1, 1);
    const history = new History();
    const controller = new InteractionController({ store, history, getWb: () => wb });
    controller.setPolicy(fixedFormPolicy({ ranges: [{ sheet: 0, r0: 0, c0: 3, r1: 3, c1: 6 }] }));
    unregisterController = registerInteractionController(store, controller);
    const handle = attachClipboard({ host, history, store, wb, onAfterCommit });

    const { transfer } = fireClipboard(host, 'copy');
    setRange(store, 0, 3, 3, 6);
    fireClipboard(host, 'paste', transfer.getData('text/plain'));
    wb.recalc();

    expect(wb.getValue({ sheet: 0, row: 3, col: 6 })).toEqual({ kind: 'number', value: 4 });
    expect(store.getState().selection.range).toEqual({ sheet: 0, r0: 0, c0: 3, r1: 3, c1: 6 });
    expect(onAfterCommit).toHaveBeenCalledTimes(1);
    expect(history.undo()).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 3, col: 6 }).kind).toBe('blank');
    handle.detach();
  });

  it('refreshes a copy snapshot after a row insertion moves its source', () => {
    seedAndMirror(store, wb, [
      { row: 1, col: 0, value: 5, formula: '=B$2' },
      { row: 1, col: 1, value: 5 },
    ]);
    expect(wb.addSheet('Target')).toBe(1);
    setRange(store, 1, 0, 1, 0);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    const { transfer } = fireClipboard(host, 'copy');
    expect(handle.getSnapshot()?.range).toEqual({ sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 });

    insertRows(store, wb, null, 1, 1);
    // The structure command shifts the live marquee, while the cached cell map
    // is refreshed by the same path the UI uses after an engine edit.
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      cells.delete(addrKey({ sheet: 0, row: 1, col: 0 }));
      cells.delete(addrKey({ sheet: 0, row: 1, col: 1 }));
      for (const cell of wb.cells(0)) {
        cells.set(addrKey(cell.addr), { value: cell.value, formula: cell.formula });
      }
      return { ...s, data: { ...s.data, cells } };
    });
    expect(store.getState().ui.copyRange).toEqual({
      sheet: 0,
      r0: 2,
      c0: 0,
      r1: 2,
      c1: 0,
    });
    expect(handle.getSnapshot()?.range).toEqual({ sheet: 0, r0: 2, c0: 0, r1: 2, c1: 0 });
    expect(handle.getSnapshot()?.cells[0]?.[0]?.formula).toBe('=B$3');

    mutators.setSheetIndex(store, 1);
    setRange(store, 4, 3, 4, 3, 1);
    fireClipboard(host, 'paste', transfer.getData('text/plain'));
    expect(wb.cellFormula({ sheet: 1, row: 4, col: 3 })).toBe('=E$3');
    handle.detach();
  });

  it('consumes a cut marquee on the first paste', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 'a' }]);
    setRange(store, 0, 0, 0, 0);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    const { transfer } = fireClipboard(host, 'cut');
    expect(store.getState().ui.copyMode).toBe('cut');

    setRange(store, 3, 3, 3, 3);
    fireClipboard(host, 'paste', transfer.getData('text/plain'));
    expect(store.getState().ui.copyRange).toBeNull();
    expect(store.getState().ui.copyMode).toBeNull();
    handle.detach();
  });

  it('moves a cut payload and restores source/destination in one undo step', () => {
    const history = new History();
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 5 }]);
    setFormat(store, 0, 0, { bold: true });
    setRange(store, 0, 0, 0, 0);
    const handle = attachClipboard({ host, history, store, wb, onAfterCommit });

    const { transfer } = fireClipboard(host, 'cut');
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 5 });
    setRange(store, 2, 2, 2, 2);
    fireClipboard(host, 'paste', transfer.getData('text/plain'));
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
    expect(wb.getValue({ sheet: 0, row: 2, col: 2 })).toEqual({ kind: 'number', value: 5 });

    expect(history.undo()).toBe(true);
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 5 });
    expect(wb.getValue({ sheet: 0, row: 2, col: 2 }).kind).toBe('blank');
    expect(formatAt(store, 0, 0)).toMatchObject({ bold: true });
    handle.detach();
  });

  it('records a multi-cell paste as one undo step when history is attached', () => {
    const history = new History();
    setRange(store, 1, 1, 1, 1);
    const handle = attachClipboard({ host, history, store, wb, onAfterCommit });

    fireClipboard(host, 'paste', 'foo\t42\r\nbar\t99');
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'text', value: 'foo' });
    expect(wb.getValue({ sheet: 0, row: 2, col: 2 })).toEqual({ kind: 'number', value: 99 });

    expect(history.undo()).toBe(true);
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 1, col: 1 }).kind).toBe('blank');
    expect(wb.getValue({ sheet: 0, row: 1, col: 2 }).kind).toBe('blank');
    expect(wb.getValue({ sheet: 0, row: 2, col: 1 }).kind).toBe('blank');
    expect(wb.getValue({ sheet: 0, row: 2, col: 2 }).kind).toBe('blank');
    expect(history.canUndo()).toBe(false);
    handle.detach();
  });

  it('paste from the current internal copy preserves source formatting by default', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 7 }]);
    setFormat(store, 0, 0, {
      bold: true,
      fill: '#fff2cc',
      numFmt: { kind: 'currency', symbol: '$', decimals: 0 },
    });
    setRange(store, 0, 0, 0, 0);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    const { transfer } = fireClipboard(host, 'copy');
    setRange(store, 2, 2, 2, 2);
    const { event } = fireClipboard(host, 'paste', transfer.getData('text/plain'));

    expect(event.defaultPrevented).toBe(true);
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 2, col: 2 })).toEqual({ kind: 'number', value: 7 });
    expect(formatAt(store, 2, 2)).toMatchObject({
      bold: true,
      fill: '#fff2cc',
      numFmt: { kind: 'currency', symbol: '$', decimals: 0 },
    });
    expect(store.getState().selection.range).toEqual({ sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 });
    expect(onAfterCommit).toHaveBeenCalledTimes(1);
    handle.detach();
  });

  it('paste with external text ignores a stale internal snapshot', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 7 }]);
    setFormat(store, 0, 0, { bold: true, fill: '#fff2cc' });
    setFormat(store, 2, 2, { italic: true, fill: '#ddebf7' });
    setRange(store, 0, 0, 0, 0);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    fireClipboard(host, 'copy');
    setRange(store, 2, 2, 2, 2);
    fireClipboard(host, 'paste', 'external');

    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 2, col: 2 })).toEqual({ kind: 'text', value: 'external' });
    expect(formatAt(store, 2, 2)).toEqual({ italic: true, fill: '#ddebf7' });
    expect(handle.getSnapshot()).toBeNull();
    handle.detach();
  });

  it('paste with empty payload is a no-op', () => {
    setRange(store, 1, 1, 1, 1);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    const { event } = fireClipboard(host, 'paste', '');
    // Empty payload short-circuits before preventDefault.
    expect(event.defaultPrevented).toBe(false);
    expect(onAfterCommit).not.toHaveBeenCalled();
    handle.detach();
  });

  it('paste is a no-op while the editor is active', () => {
    setRange(store, 1, 1, 1, 1);
    store.setState((s) => ({
      ...s,
      ui: { ...s.ui, editor: { kind: 'edit', raw: '', caret: 0 } },
    }));
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    fireClipboard(host, 'paste', 'foo\t42');
    expect(onAfterCommit).not.toHaveBeenCalled();
    expect(wb.getValue({ sheet: 0, row: 1, col: 1 }).kind).toBe('blank');
    handle.detach();
  });

  it('detach removes listeners so subsequent events do not fire handlers', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 1 }]);
    setRange(store, 0, 0, 0, 0);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });
    handle.detach();

    const { event } = fireClipboard(host, 'copy');
    expect(event.defaultPrevented).toBe(false);
    expect(handle.getSnapshot()).toBeNull();
  });

  it('copy without clipboardData on the event short-circuits gracefully', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 1 }]);
    setRange(store, 0, 0, 0, 0);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    // Construct a copy event without clipboardData.
    const event = new ClipboardEvent('copy', { bubbles: true, cancelable: true });
    host.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(false);
    handle.detach();
  });
});
