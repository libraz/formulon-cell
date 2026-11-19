import { FormulonError, StatusCode, type Workbook } from '@libraz/formulon';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  type SheetActionContext,
  setSheetTabColorAction,
} from '../../../src/commands/sheet-actions.js';
import { syncSheetPrToEngine } from '../../../src/engine/sheet-pr-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

/** Makes every `<sheetPr>` write refuse; returns the tab colour setter spy. */
const refuse = (wb: WorkbookHandle) => {
  const sheet = (wb as unknown as { wb: Workbook }).wb.sheet(0);
  const ooxml = Object.getPrototypeOf(sheet.ooxml) as {
    setSheetPrXml: (xml: string | null) => void;
  };
  const sheetProto = Object.getPrototypeOf(sheet) as { setTabColor: (color: unknown) => void };
  const refused = () => {
    throw new FormulonError(StatusCode.InvalidArgument, 'refused', '');
  };
  vi.spyOn(ooxml, 'setSheetPrXml').mockImplementation(refused);
  return vi.spyOn(sheetProto, 'setTabColor').mockImplementation(refused);
};

describe('sheetPr write refused by the engine', () => {
  const handles: WorkbookHandle[] = [];
  afterEach(() => {
    vi.restoreAllMocks();
    for (const h of handles.splice(0)) h.dispose();
  });

  const setup = async () => {
    const workbook = await WorkbookHandle.createDefault();
    handles.push(workbook);
    const store = createSpreadsheetStore();
    workbook.attachStore(store);
    const history = new History();
    workbook.attachHistory(history);
    const onRefused = vi.fn();
    workbook.subscribeWriteRefusals(onRefused);
    const ctx: SheetActionContext = { store, workbook, history, origin: 'contextMenu' };
    return { store, workbook, history, ctx, onRefused };
  };

  it('reverts the tab colour to the engine value, notifies once and records no undo step', async () => {
    const { store, workbook, history, ctx, onRefused } = await setup();
    const spy = refuse(workbook);
    setSheetTabColorAction(ctx, 0, '#ff0000');
    expect(spy).toHaveBeenCalled();
    expect(store.getState().layout.sheetTabColors.has(0)).toBe(false);
    expect(workbook.getSheetPr(0)?.tabColor).toBeNull();
    expect(onRefused).toHaveBeenCalledTimes(1);
    expect(onRefused).toHaveBeenCalledWith({ kind: 'sheetPr', sheet: 0 });
    expect(history.canUndo()).toBe(false);
  });

  it('reverts a refused tab colour change to the colour the engine holds', async () => {
    const { store, workbook, history, ctx, onRefused } = await setup();
    expect(setSheetTabColorAction(ctx, 0, '#00b050')).toBe(true);
    expect(onRefused).not.toHaveBeenCalled();
    const spy = refuse(workbook);
    setSheetTabColorAction(ctx, 0, null);
    expect(store.getState().layout.sheetTabColors.get(0)).toBe('#00b050');
    expect(onRefused).toHaveBeenCalledTimes(1);
    spy.mockRestore();
    expect(history.undo()).toBe(true);
    expect(store.getState().layout.sheetTabColors.has(0)).toBe(false);
    expect(history.canUndo()).toBe(false);
  });

  it('reverts the outline summary position and reports the refusal from the sync', async () => {
    const { store, workbook, onRefused } = await setup();
    refuse(workbook);
    mutators.setOutlineSummary(store, 0, { below: false });
    expect(store.getState().layout.outlineSummary?.get(0)).toBeUndefined();
    expect(workbook.getSheetPr(0)?.summaryBelow).toBe(true);
    expect(onRefused).toHaveBeenCalledTimes(1);
    expect(syncSheetPrToEngine(workbook, store, 0)).toEqual({ ok: true });
  });

  it('reverts and notifies once when an undo replay is refused', async () => {
    const { store, workbook, history, ctx, onRefused } = await setup();
    expect(setSheetTabColorAction(ctx, 0, '#c00000')).toBe(true);
    refuse(workbook);
    expect(history.undo()).toBe(true);
    expect(store.getState().layout.sheetTabColors.get(0)).toBe('#c00000');
    expect(workbook.getSheetPr(0)?.tabColor).toEqual({ rgb: 'FFC00000' });
    expect(onRefused).toHaveBeenCalledTimes(1);
  });

  it('reverts and notifies once when a redo replay is refused', async () => {
    const { store, workbook, history, ctx, onRefused } = await setup();
    expect(setSheetTabColorAction(ctx, 0, '#c00000')).toBe(true);
    expect(history.undo()).toBe(true);
    refuse(workbook);
    expect(history.redo()).toBe(true);
    expect(store.getState().layout.sheetTabColors.has(0)).toBe(false);
    expect(workbook.getSheetPr(0)?.tabColor).toBeNull();
    expect(onRefused).toHaveBeenCalledTimes(1);
  });

  it('keeps the success path: store, engine and one undo step agree', async () => {
    const { store, workbook, history, ctx, onRefused } = await setup();
    expect(setSheetTabColorAction(ctx, 0, '#c00000')).toBe(true);
    expect(store.getState().layout.sheetTabColors.get(0)).toBe('#c00000');
    expect(workbook.getSheetPr(0)?.tabColor).toEqual({ rgb: 'FFC00000' });
    expect(onRefused).not.toHaveBeenCalled();
    expect(history.undo()).toBe(true);
    expect(workbook.getSheetPr(0)?.tabColor).toBeNull();
  });

  it('takes the tab colour back when the outline part of the same write is refused', async () => {
    const { workbook } = await setup();
    const ooxml = Object.getPrototypeOf(
      (workbook as unknown as { wb: Workbook }).wb.sheet(0).ooxml,
    ) as {
      setSheetPrXml: (xml: string | null) => void;
    };
    vi.spyOn(ooxml, 'setSheetPrXml').mockImplementation(() => {
      throw new FormulonError(StatusCode.InvalidArgument, 'refused', '');
    });
    const written = workbook.setSheetPr(0, { tabColor: { rgb: 'FFFF0000' }, summaryBelow: false });
    expect(written.ok).toBe(false);
    expect(workbook.getSheetPr(0)?.tabColor).toBeNull();
    expect(workbook.getSheetPr(0)?.summaryBelow).toBe(true);
  });
});
