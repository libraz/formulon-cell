import { afterEach, describe, expect, it } from 'vitest';
import {
  builtInProtectionDecision,
  canExecuteBuiltIn,
  cellShiftChoices,
} from '../../../src/commands/built-in-command-policy.js';
import type { InteractionPolicy } from '../../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { insertCellsAction, insertRowsAction } from '../../../src/commands/row-col-actions.js';
import { createDefaultDynamicDropdownsCtx } from '../../../src/mount/dynamic-dropdowns-defaults.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

const CELL = { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 };
const ROW = { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 16_383 };

const menuOf = (attr: 'cellInsert' | 'cellDelete', actions: readonly string[]): HTMLElement => {
  const menu = document.createElement('div');
  for (const action of actions) {
    const button = document.createElement('button');
    button.dataset[attr] = action;
    menu.appendChild(button);
  }
  return menu;
};

const enabled = (menu: HTMLElement, attr: 'cellInsert' | 'cellDelete'): string[] =>
  [...menu.querySelectorAll<HTMLButtonElement>('button')]
    .filter((button) => !button.disabled)
    .map((button) => button.dataset[attr] ?? '');

describe('row and column insert and delete on a protected sheet', () => {
  let sheet: MountedStubSheet | null = null;

  afterEach(() => {
    sheet?.dispose();
    sheet = null;
    document.body.replaceChildren();
  });

  const setup = async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    const { instance, workbook } = sheet;
    for (let row = 0; row < 4; row += 1) workbook.setNumber({ sheet: 0, row, col: 0 }, row + 1);
    mutators.replaceCells(instance.store, workbook.cells(0));
    const ctx = createDefaultDynamicDropdownsCtx(instance);
    const value = (row: number) => workbook.getValue({ sheet: 0, row, col: 0 });
    return { instance, ctx, value };
  };

  it('opens no cell dialog and offers nothing without an insert option', async () => {
    const { instance, ctx } = await setup();
    setProtectedSheet(instance.store, 0, true);
    const menu = menuOf('cellInsert', ['cells', 'rows', 'cols']);
    mutators.setRange(instance.store, CELL);
    ctx.updateCellInsertMenu(menu);
    expect(enabled(menu, 'cellInsert')).toEqual([]);
    ctx.applyCellInsertAction('cells');
    expect(document.querySelector('.fc-cellshift')).toBeNull();
  });

  it('with Insert rows only, offers Rows; the Home Cells item stays off, the context menu dialog offers Entire row', async () => {
    const { instance, ctx, value } = await setup();
    setProtectedSheet(instance.store, 0, true, { permissions: { insertRows: true } });
    const menu = menuOf('cellInsert', ['cells', 'rows', 'cols']);

    mutators.setRange(instance.store, CELL);
    ctx.updateCellInsertMenu(menu);
    expect(enabled(menu, 'cellInsert')).toEqual(['rows']);
    ctx.applyCellInsertAction('cells');
    expect(document.querySelector('.fc-cellshift')).toBeNull();
    // The context menu's Insert… opens the dialog for its Entire row choice.
    expect(canExecuteBuiltIn(instance.store, 'insertCells', 'contextMenu').allowed).toBe(true);
    expect(cellShiftChoices(instance.store, 'insert', 'contextMenu')).toEqual({
      shift: false,
      rows: true,
      cols: false,
    });
    ctx.applyCellInsertAction('cols');
    expect(value(1)).toEqual({ kind: 'number', value: 2 });

    mutators.setRange(instance.store, ROW);
    ctx.updateCellInsertMenu(menu);
    expect(enabled(menu, 'cellInsert')).toContain('rows');
    expect(enabled(menu, 'cellInsert')).not.toContain('cols');
    ctx.applyCellInsertAction('rows');
    expect(value(1)).toEqual({ kind: 'blank' });
    expect(value(2)).toEqual({ kind: 'number', value: 2 });
  });

  it('with Delete rows only, offers Rows and, on the Mac desktop, deletes all-locked rows', async () => {
    const { instance, ctx, value } = await setup();
    // The Mac desktop ignores the lock of the rows it deletes; Windows requires them unlocked.
    instance.host.dataset.fcPlatform = 'mac';
    setProtectedSheet(instance.store, 0, true, { permissions: { deleteRows: true } });
    const menu = menuOf('cellDelete', ['cells', 'rows', 'row', 'cols', 'col']);

    mutators.setRange(instance.store, CELL);
    ctx.updateCellDeleteMenu(menu);
    expect(enabled(menu, 'cellDelete')).toEqual(['rows', 'row']);
    ctx.applyCellDeleteAction('cells');
    expect(document.querySelector('.fc-cellshift')).toBeNull();

    // Every cell is locked; the option alone lets the rows go.
    mutators.setRange(instance.store, ROW);
    ctx.applyCellDeleteAction('rows');
    expect(value(1)).toEqual({ kind: 'number', value: 3 });
  });

  it('on Windows keeps rows holding a locked cell, with Delete rows on', async () => {
    const { instance, ctx, value } = await setup();
    setProtectedSheet(instance.store, 0, true, { permissions: { deleteRows: true } });
    mutators.setRange(instance.store, ROW);
    ctx.applyCellDeleteAction('rows');
    expect(value(1)).toEqual({ kind: 'number', value: 2 });
  });

  it('greys the table row and column inserts under any protection', async () => {
    const { instance } = await setup();
    const { store } = instance;
    setProtectedSheet(store, 0, true, {
      permissions: { insertRows: true, deleteRows: true, insertColumns: true, deleteColumns: true },
    });
    mutators.setRange(store, CELL);
    for (const id of ['insertTableRows', 'insertTableColumns'])
      expect(builtInProtectionDecision(store, id).allowed, id).toBe(false);
    // Insert… / Delete… stay available for their Entire row / Entire column choices.
    for (const id of ['insertCells', 'deleteCells', 'rowInsertAbove', 'rowDelete', 'colInsertLeft'])
      expect(builtInProtectionDecision(store, id).allowed, id).toBe(true);
    setProtectedSheet(store, 0, true);
    for (const id of ['insertCells', 'deleteCells', 'insertTableRows', 'rowInsertAbove'])
      expect(builtInProtectionDecision(store, id).allowed, id).toBe(false);
  });

  for (const [label, policy] of [
    ['without a policy', undefined],
    ['under an allow-all policy', { defaultOperation: 'allow' } as InteractionPolicy],
  ] as const) {
    it(`refuses a cell shift in the authorizer, whole rows going by their option ${label}`, async () => {
      const { instance, value } = await setup();
      instance.commands.setPolicy(policy);
      setProtectedSheet(instance.store, 0, true, { permissions: { insertRows: true } });
      const ctx = {
        store: instance.store,
        workbook: instance.workbook,
        history: instance.history,
        origin: 'ribbon',
        commandId: 'insertCells',
      } as const;
      expect(insertCellsAction(ctx, CELL, 'down')).toBe(false);
      expect(value(1)).toEqual({ kind: 'number', value: 2 });
      expect(insertRowsAction({ ...ctx, commandId: 'rowInsertAbove' }, 1)).toBe(true);
      expect(value(2)).toEqual({ kind: 'number', value: 2 });
    });
  }
});
