import { afterEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createFormulasDropdownDefaults } from '../../../src/mount/dynamic-dropdown-defaults/formulas.js';
import { mutators } from '../../../src/store/store.js';
import { showCreateNamesDialog } from '../../../src/toolbar/dialogs/create-names.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

const settle = () => new Promise((resolve) => setTimeout(resolve, 0));
const edge = (name: string) =>
  document.querySelector<HTMLInputElement>(
    `.fc-createnames [data-fc-edge="${name}"]`,
  ) as HTMLInputElement;
const ok = () =>
  edge('top')
    ?.closest('[role="dialog"]')
    ?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary') ?? null;

const labels = {
  title: 'Create Names',
  createIn: 'Create names in:',
  top: 'Top row',
  left: 'Left column',
  bottom: 'Bottom row',
  right: 'Right column',
  ok: 'OK',
  cancel: 'Cancel',
};

afterEach(() => {
  document.body.replaceChildren();
});

describe('Create Names dialog', () => {
  it('lists the four edges, greys OK while none is checked and resolves with the checked ones', async () => {
    const pending = showCreateNamesDialog({
      labels,
      initial: { top: false, left: false, bottom: false, right: false },
    });
    expect(
      [...document.querySelectorAll('.fc-createnames__edge')].map((e) => e.textContent?.trim()),
    ).toEqual(['Top row', 'Left column', 'Bottom row', 'Right column']);
    expect(ok()?.disabled).toBe(true);
    for (const name of ['top', 'right']) {
      edge(name).checked = true;
      edge(name).dispatchEvent(new Event('change'));
    }
    expect(ok()?.disabled).toBe(false);
    ok()?.click();
    await expect(pending).resolves.toEqual({ top: true, left: false, bottom: false, right: true });
  });
});

describe('Create from Selection', () => {
  let sheet: MountedStubSheet | null = null;
  afterEach(() => {
    sheet?.dispose();
    sheet = null;
  });

  it('creates names from the top row and the left column at once, each leaving out the other', async () => {
    sheet = await mountStubSheet({ locale: 'en', workbook: await WorkbookHandle.createDefault() });
    const { instance } = sheet;
    const put = (row: number, col: number, value: string | number) =>
      typeof value === 'number'
        ? instance.workbook.setNumber({ sheet: 0, row, col }, value)
        : instance.workbook.setText({ sheet: 0, row, col }, value);
    put(0, 1, 'Spring');
    put(0, 2, 'Summer');
    put(1, 0, 'East');
    put(2, 0, 'West');
    put(1, 1, 10);
    put(1, 2, 20);
    put(2, 1, 30);
    put(2, 2, 40);
    mutators.replaceCells(instance.store, instance.workbook.cells(0));
    mutators.setRange(instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 });

    const done =
      createFormulasDropdownDefaults(instance).applyDefinedNameAction('create-from-selection');
    await settle();
    // Both label edges hold text, so they start checked.
    expect([edge('top').checked, edge('left').checked, edge('bottom').checked]).toEqual([
      true,
      true,
      false,
    ]);
    ok()?.click();
    await done;
    const names = Object.fromEntries(
      [...instance.workbook.definedNames()].map((entry) => [entry.name, entry.formula]),
    );
    expect(names).toEqual({
      Spring: 'Sheet1!$B$2:$B$3',
      Summer: 'Sheet1!$C$2:$C$3',
      East: 'Sheet1!$B$2:$C$2',
      West: 'Sheet1!$B$3:$C$3',
    });
    instance.history.undo();
    expect([...instance.workbook.definedNames()]).toEqual([]);
  });

  // Provisional until measured: the shared corner cell decides neither edge.
  it('starts with only the top row checked over a header row of numbers', async () => {
    sheet = await mountStubSheet({ locale: 'en', workbook: await WorkbookHandle.createDefault() });
    const { instance } = sheet;
    instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'Qty');
    instance.workbook.setText({ sheet: 0, row: 0, col: 1 }, 'Price');
    for (const row of [1, 2]) {
      instance.workbook.setNumber({ sheet: 0, row, col: 0 }, row * 2 + 1);
      instance.workbook.setNumber({ sheet: 0, row, col: 1 }, row * 10);
    }
    mutators.replaceCells(instance.store, instance.workbook.cells(0));
    mutators.setRange(instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 });
    void createFormulasDropdownDefaults(instance).applyDefinedNameAction('create-from-selection');
    await settle();
    expect([edge('top').checked, edge('left').checked]).toEqual([true, false]);
  });

  // Provisional: every name the workbook already holds is asked about before any is written.
  it('asks before replacing a name the workbook holds: No keeps it, Yes replaces it', async () => {
    sheet = await mountStubSheet({ locale: 'en', workbook: await WorkbookHandle.createDefault() });
    const { instance } = sheet;
    instance.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'Qty');
    instance.workbook.setText({ sheet: 0, row: 0, col: 1 }, 'Price');
    mutators.replaceCells(instance.store, instance.workbook.cells(0));
    mutators.setRange(instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 });
    instance.workbook.setDefinedNameEntry('QTY', '=Sheet1!$Z$1');
    const names = () =>
      Object.fromEntries(
        [...instance.workbook.definedNames()].map((entry) => [entry.name, entry.formula]),
      );
    /** Runs Create from Selection with Top row, answering each replace alert in turn. */
    const run = async (...answers: ('Yes' | 'No')[]): Promise<(string | null | undefined)[]> => {
      const done =
        createFormulasDropdownDefaults(instance).applyDefinedNameAction('create-from-selection');
      await settle();
      ok()?.click();
      const messages = [];
      for (const answer of answers) {
        await settle();
        const alert = document.querySelector('[role="alertdialog"]');
        messages.push(alert?.querySelector('.fc-tb__dlg__message')?.textContent);
        [...(alert?.querySelectorAll('button') ?? [])]
          .find((button) => button.textContent === answer)
          ?.click();
      }
      await done;
      return messages;
    };

    expect(await run('No')).toEqual(["Replace existing definition of 'Qty'?"]);
    expect(names()).toEqual({ QTY: 'Sheet1!$Z$1', Price: 'Sheet1!$B$2:$B$3' });
    expect(await run('Yes', 'No')).toEqual([
      "Replace existing definition of 'Qty'?",
      "Replace existing definition of 'Price'?",
    ]);
    expect(names()).toEqual({ Qty: 'Sheet1!$A$2:$A$3', Price: 'Sheet1!$B$2:$B$3' });
  });
});
