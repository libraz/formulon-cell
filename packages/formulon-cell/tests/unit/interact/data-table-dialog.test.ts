import { afterEach, describe, expect, it } from 'vitest';
import { clearSelectedContents } from '../../../src/commands/clear-contents.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { openDataTableDialog } from '../../../src/interact/data-table-dialog.js';
import { mutators } from '../../../src/store/store.js';
import { projectMacRibbonState } from '../../../src/toolbar/ribbon/mac/model.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

const q = <T extends HTMLElement>(selector: string): T => {
  const el = document.querySelector<T>(selector);
  if (!el) throw new Error(`missing ${selector}`);
  return el;
};
const field = (name: 'row' | 'col'): HTMLInputElement =>
  q<HTMLInputElement>(`.fc-datatable [data-data-table-input="${name}"]`);
const ok = (): void => q<HTMLButtonElement>('.fc-datatable .fc-fmtdlg__btn--primary').click();
const status = (): string => q('.fc-datatable__status').textContent ?? '';

describe.skipIf(!canLoadWasm())('Data Table dialog', () => {
  let sheet: MountedStubSheet | null = null;
  afterEach(() => {
    sheet?.dispose();
    sheet = null;
  });

  const mountWithModel = async () => {
    const workbook = await WorkbookHandle.createDefault();
    workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 5);
    for (const i of [1, 2, 3]) workbook.setNumber({ sheet: 0, row: 1 + i, col: 1 }, i);
    workbook.setFormula({ sheet: 0, row: 1, col: 2 }, '=A1*2');
    sheet = await mountStubSheet({ workbook, locale: 'en' });
    const { store } = sheet.instance;
    mutators.setRange(store, { sheet: 0, r0: 1, c0: 1, r1: 4, c1: 2 });
    return { workbook, instance: sheet.instance };
  };

  it('creates a one-variable table over the selection and undoes it', async () => {
    const { workbook, instance } = await mountWithModel();
    openDataTableDialog(instance);
    field('col').value = 'A1';
    ok();
    expect(document.querySelector<HTMLElement>('.fc-datatable')?.hidden).toBe(true);
    expect(workbook.dataTableAt({ sheet: 0, row: 3, col: 2 })?.range).toEqual({
      sheet: 0,
      r0: 2,
      c0: 2,
      r1: 4,
      c1: 2,
    });
    expect(instance.store.getState().data.cells.get('0:4:2')?.value).toEqual({
      kind: 'number',
      value: 6,
    });
    instance.history.undo();
    expect(workbook.dataTables(0)).toEqual([]);
    expect(instance.store.getState().data.cells.get('0:4:2')?.value ?? { kind: 'blank' }).toEqual({
      kind: 'blank',
    });
  });

  it('asks for an input cell and keeps the dialog open', async () => {
    const { workbook, instance } = await mountWithModel();
    openDataTableDialog(instance);
    ok();
    expect(status()).toBe(en.dataTableDialog.invalidInput);
    field('row').value = 'not a cell';
    ok();
    expect(status()).toBe(en.dataTableDialog.invalidInput);
    expect(workbook.dataTables(0)).toEqual([]);
  });

  it('refuses a selection without a formula row and value column', async () => {
    const { instance } = await mountWithModel();
    mutators.setRange(instance.store, { sheet: 0, r0: 1, c0: 1, r1: 4, c1: 1 });
    openDataTableDialog(instance);
    field('col').value = 'A1';
    ok();
    expect(status()).toBe(en.dataTableDialog.invalidRange);
  });

  it('refuses an input cell inside the results', async () => {
    const { instance } = await mountWithModel();
    openDataTableDialog(instance);
    field('col').value = 'C4';
    ok();
    expect(status()).toBe(en.dataTableDialog.invalidInput);
  });

  it('offers the menu item only on an engine that saves data tables', async () => {
    const { instance } = await mountWithModel();
    const menu = document.createElement('div');
    menu.innerHTML = '<button data-ribbon-command="mac.data.dataTable"></button>';
    const item = () => menu.querySelector('button')?.getAttribute('aria-disabled') === 'true';
    projectMacRibbonState(menu, instance);
    expect(item()).toBe(false);
    sheet?.dispose();
    sheet = await mountStubSheet();
    projectMacRibbonState(menu, sheet.instance);
    expect(item()).toBe(true);
  });

  it('refuses edits to a table loaded from a file with the desktop alert', async () => {
    const source = await WorkbookHandle.createDefault();
    source.setNumber({ sheet: 0, row: 0, col: 0 }, 5);
    for (const i of [1, 2, 3]) source.setNumber({ sheet: 0, row: 1 + i, col: 1 }, i);
    source.setFormula({ sheet: 0, row: 1, col: 2 }, '=A1*2');
    source.addDataTable({
      range: { sheet: 0, r0: 2, c0: 2, r1: 4, c1: 2 },
      rowInput: null,
      colInput: { sheet: 0, row: 0, col: 0 },
    });
    const bytes = source.save();
    source.dispose();
    const workbook = await WorkbookHandle.loadBytes(bytes);
    sheet = await mountStubSheet({ workbook, locale: 'en' });
    const { instance } = sheet;
    expect(instance.store.getState().data.cells.get('0:3:2')?.value).toEqual({
      kind: 'number',
      value: 4,
    });
    mutators.setRange(instance.store, { sheet: 0, r0: 3, c0: 2, r1: 3, c1: 2 });
    expect(clearSelectedContents({ store: instance.store, workbook }).status).toBe('rejected');
    expect(document.body.textContent).toContain(en.dataTableDialog.cannotChange);
  });

  it('shows a result cell as a read-only {=TABLE()} array formula', async () => {
    const { workbook, instance } = await mountWithModel();
    workbook.addDataTable({
      range: { sheet: 0, r0: 2, c0: 2, r1: 4, c1: 2 },
      rowInput: null,
      colInput: { sheet: 0, row: 0, col: 0 },
    });
    mutators.replaceCells(instance.store, workbook.cells(0));
    const fx = q<HTMLTextAreaElement>('.fc-host__formulabar-input');
    mutators.setActive(instance.store, { sheet: 0, row: 3, col: 2 });
    expect(fx.value).toBe('{=TABLE(,A1)}');
    expect(fx.readOnly).toBe(true);
    fx.dispatchEvent(new FocusEvent('focus'));
    expect(q('.fc-host__formulabar').dataset.fcEditing).not.toBe('1');
    mutators.setActive(instance.store, { sheet: 0, row: 1, col: 2 });
    expect(fx.value).toBe('=A1*2');
    expect(fx.readOnly).toBe(false);
  });

  it('restores what the result block held before on undo', async () => {
    const { workbook, instance } = await mountWithModel();
    workbook.setText({ sheet: 0, row: 3, col: 2 }, 'old');
    workbook.setFormula({ sheet: 0, row: 4, col: 2 }, '=B5+100');
    mutators.replaceCells(instance.store, workbook.cells(0));
    instance.history.clear();
    openDataTableDialog(instance);
    field('col').value = 'A1';
    ok();
    expect(workbook.dataTables(0)).toHaveLength(1);
    instance.history.undo();
    expect(workbook.dataTables(0)).toEqual([]);
    expect(workbook.getValue({ sheet: 0, row: 3, col: 2 })).toEqual({ kind: 'text', value: 'old' });
    expect(workbook.cellFormula({ sheet: 0, row: 4, col: 2 })).toBe('=B5+100');
    instance.history.redo();
    expect(workbook.dataTables(0)).toHaveLength(1);
  });

  it('refuses a protected sheet or a locked result cell', async () => {
    const { workbook, instance } = await mountWithModel();
    mutators.setCellFormat(instance.store, { sheet: 0, row: 3, col: 2 }, { locked: true });
    instance.store.setState((s) => ({
      ...s,
      protection: {
        ...s.protection,
        protectedSheets: new Map([[0, { password: null } as never]]),
      },
    }));
    openDataTableDialog(instance);
    field('col').value = 'A1';
    ok();
    expect(workbook.dataTables(0)).toEqual([]);
    expect(status()).toBe(en.dataTableDialog.rejected);
  });
});
