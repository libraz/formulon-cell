import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  type AllowedRangeUnlockUi,
  registerAllowedRangeUnlockUi,
  setProtectedSheet,
} from '../../../src/commands/protection.js';
import { legacySheetPasswordHash } from '../../../src/engine/protection-password.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachBorderDraw } from '../../../src/interact/border-draw.js';
import { attachKeyboard } from '../../../src/interact/keyboard.js';
import { createDataDropdownDefaults } from '../../../src/mount/dynamic-dropdown-defaults/data.js';
import { createEditingDropdownDefaults } from '../../../src/mount/dynamic-dropdown-defaults/editing.js';
import { createInsertDropdownDefaults } from '../../../src/mount/dynamic-dropdown-defaults/insert.js';
import { cellRect, layoutForView } from '../../../src/render/geometry.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';
import { waitFor } from '../mount/toolbar/fixtures.js';

const PASSWORD = 'pw';
const settle = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 0));
const area = (r0: number, c0: number, r1: number, c1: number): Range => ({
  sheet: 0,
  r0,
  c0,
  r1,
  c1,
});
const A1 = { sheet: 0, row: 0, col: 0 };

const guard = (store: SpreadsheetStore, range: Range, answers: (string | null)[] = [PASSWORD]) => {
  mutators.addAllowedEditRange(store, {
    title: 'Guarded',
    range,
    legacyPasswordHash: legacySheetPasswordHash(PASSWORD),
  });
  setProtectedSheet(store, 0, true);
  const ui: AllowedRangeUnlockUi & {
    askPassword: ReturnType<typeof vi.fn>;
    reportIncorrect: ReturnType<typeof vi.fn>;
  } = {
    askPassword: vi.fn(async () => answers.shift() ?? null),
    reportIncorrect: vi.fn(async () => {}),
  };
  registerAllowedRangeUnlockUi(store, ui);
  return ui;
};

const select = (store: SpreadsheetStore, range: Range): void => {
  mutators.setActive(store, { sheet: 0, row: range.r0, col: range.c0 });
  mutators.setRange(store, range);
};

describe('on a mounted sheet', () => {
  let sheet: MountedStubSheet;
  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
  });
  afterEach(() => {
    sheet.dispose();
    document.body.innerHTML = '';
  });

  it('formula bar: the prompt moving focus does not commit twice, and Enter still moves down', async () => {
    const { store } = sheet.instance;
    select(store, area(0, 0, 0, 0));
    mutators.addAllowedEditRange(store, {
      title: 'Guarded',
      range: area(0, 0, 0, 0),
      legacyPasswordHash: legacySheetPasswordHash(PASSWORD),
    });
    setProtectedSheet(store, 0, true);
    // The mount's own prompt (a real dialog) stays registered.
    const fx = sheet.host.querySelector('.fc-host__formulabar-input') as HTMLTextAreaElement;
    fx.focus();
    fx.dispatchEvent(new FocusEvent('focus'));
    fx.value = '7';
    fx.dispatchEvent(new Event('input'));
    fx.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', cancelable: true }));
    await waitFor(() => document.querySelectorAll('.fc-tb__dlg').length === 1);

    // Opening the dialog moves focus away from the formula bar.
    fx.dispatchEvent(new FocusEvent('blur'));
    await settle();
    expect(document.querySelectorAll('.fc-tb__dlg')).toHaveLength(1);

    const input = document.querySelector('.fc-tb__dlg input') as HTMLInputElement;
    input.value = PASSWORD;
    (document.querySelector('.fc-tb__dlg .fc-fmtdlg__btn--primary') as HTMLButtonElement).click();
    await waitFor(() => sheet.workbook.getValue(A1).kind === 'number');
    expect(store.getState().selection.active).toMatchObject({ row: 1, col: 0 });
    expect(document.querySelectorAll('.fc-tb__dlg')).toHaveLength(0);
  });

  it('formula bar: a cancelled prompt is not asked again by the blur', async () => {
    const { store } = sheet.instance;
    select(store, area(0, 0, 0, 0));
    const ui = guard(store, area(0, 0, 0, 0), [null]);
    const fx = sheet.host.querySelector('.fc-host__formulabar-input') as HTMLTextAreaElement;
    fx.focus();
    fx.dispatchEvent(new FocusEvent('focus'));
    fx.value = '7';
    fx.dispatchEvent(new Event('input'));
    fx.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', cancelable: true }));
    fx.dispatchEvent(new FocusEvent('blur'));
    await settle();
    await settle();
    expect(ui.askPassword).toHaveBeenCalledTimes(1);
    expect(sheet.workbook.getValue(A1).kind).toBe('blank');
  });

  it('Clear Contents asks for a guarded range in an extra selection area', async () => {
    const { store } = sheet.instance;
    sheet.workbook.setNumber({ sheet: 0, row: 4, col: 4 }, 3);
    sheet.workbook.recalc();
    mutators.replaceCells(store, sheet.workbook.cells(0));
    select(store, area(0, 0, 0, 0));
    store.setState((s) => ({
      ...s,
      selection: { ...s.selection, extraRanges: [area(4, 4, 4, 4)] },
    }));
    const ui = guard(store, area(4, 4, 4, 4), [PASSWORD]);
    mutators.setCellFormat(store, A1, { locked: false });
    createEditingDropdownDefaults(sheet.instance).applyClearAction('contents');
    await settle();
    expect(ui.askPassword).toHaveBeenCalledTimes(1);
  });

  it('Insert Symbol cannot write a locked cell, and asks for a guarded range', async () => {
    const { store } = sheet.instance;
    select(store, area(0, 0, 0, 0));
    setProtectedSheet(store, 0, true);
    createInsertDropdownDefaults(sheet.instance).applySymbolAction('©');
    await settle();
    expect(sheet.workbook.getValue(A1).kind).toBe('blank');

    setProtectedSheet(store, 0, false);
    const ui = guard(store, area(0, 0, 0, 0));
    createInsertDropdownDefaults(sheet.instance).applySymbolAction('©');
    expect(sheet.workbook.getValue(A1).kind).toBe('blank');
    await settle();
    expect(ui.askPassword).toHaveBeenCalledTimes(1);
    expect(sheet.workbook.getValue(A1)).toEqual({ kind: 'text', value: '©' });
  });

  it('Text to Columns prompts only for the cells it really writes', async () => {
    const { store } = sheet.instance;
    sheet.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'a,b');
    sheet.workbook.recalc();
    mutators.replaceCells(store, sheet.workbook.cells(0));
    select(store, area(0, 0, 0, 0));
    // Source cell and its one output cell are guarded; nothing further right is touched.
    const ui = guard(store, area(0, 0, 0, 1));
    createDataDropdownDefaults(sheet.instance).splitTextToColumns(',');
    await settle();
    expect(ui.askPassword).toHaveBeenCalledTimes(1);
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'text',
      value: 'b',
    });
  });
});

describe('border pen', () => {
  it('a prompt ends the stroke instead of queueing more prompts', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const ui = guard(store, area(0, 0, 3, 3), [null]);
    const pen = attachBorderDraw({ host, store });
    pen.activate('draw');
    const s = store.getState();
    const at = (row: number, type: string) => {
      const r = cellRect(layoutForView(s), s.viewport, row, 0);
      host.dispatchEvent(
        new MouseEvent(type, {
          clientX: r.x + r.w / 2,
          clientY: r.y + r.h - 1,
          bubbles: true,
          cancelable: true,
        }),
      );
    };
    at(0, 'pointerdown');
    await settle();
    at(1, 'pointermove');
    await settle();
    expect(ui.askPassword).toHaveBeenCalledTimes(1);
    pen.detach();
    document.body.innerHTML = '';
  });
});

describe('Backspace without an interaction controller', () => {
  it('does not clear a locked cell and asks for a guarded range first', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    wb.setNumber(A1, 5);
    wb.recalc();
    const onBeginEdit = vi.fn();
    const detach = attachKeyboard({
      host,
      store,
      wb,
      history: new History(),
      onBeginEdit,
      onClearActive: vi.fn(),
      onGoTo: vi.fn(),
      onSwitchSheet: vi.fn(),
      onPasteSpecial: vi.fn(),
    });
    const press = () =>
      host.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'Backspace', bubbles: true, cancelable: true }),
      );
    setProtectedSheet(store, 0, true);
    press();
    await settle();
    expect(wb.getValue(A1)).toEqual({ kind: 'number', value: 5 });

    setProtectedSheet(store, 0, false);
    const ui = guard(store, area(0, 0, 0, 0));
    press();
    expect(wb.getValue(A1)).toEqual({ kind: 'number', value: 5 });
    await settle();
    expect(ui.askPassword).toHaveBeenCalledTimes(1);
    expect(wb.getValue(A1).kind).toBe('blank');
    detach();
    document.body.innerHTML = '';
  });
});
