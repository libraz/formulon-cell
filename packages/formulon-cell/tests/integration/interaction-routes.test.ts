import { afterEach, describe, expect, it } from 'vitest';
import { fixedFormPolicy, viewerPolicy } from '../../src/commands/interaction-policy.js';
import { mutators } from '../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

const addr = (row: number, col: number) => ({ sheet: 0, row, col });

const paste = (host: HTMLElement, text: string): ClipboardEvent => {
  const transfer = new DataTransfer();
  transfer.setData('text/plain', text);
  const event = new ClipboardEvent('paste', {
    clipboardData: transfer as unknown as DataTransfer,
    bubbles: true,
    cancelable: true,
  });
  host.dispatchEvent(event);
  return event;
};

const key = (host: HTMLElement, keyName: string, init: KeyboardEventInit = {}): KeyboardEvent => {
  const event = new KeyboardEvent('keydown', {
    key: keyName,
    bubbles: true,
    cancelable: true,
    ...init,
  });
  host.dispatchEvent(event);
  return event;
};

describe('restricted interaction routes', () => {
  let sheet: MountedStubSheet;

  afterEach(() => sheet?.dispose());

  it('viewer keeps values unchanged through typing, clear keys, formula bar, paste, and pointer selection', async () => {
    sheet = await mountStubSheet({
      policy: viewerPolicy(),
      features: { clipboard: true, formulaBar: true, shortcuts: true },
    });
    const { host, instance, workbook } = sheet;
    const target = addr(0, 0);
    workbook.setText(target, 'original');
    mutators.replaceCells(instance.store, workbook.cells(0));
    mutators.setActive(instance.store, target);
    expect(instance.commands.policy).toEqual(viewerPolicy());
    expect(workbook.getValue(target)).toEqual({ kind: 'text', value: 'original' });

    key(host, 'x');
    expect(host.querySelector('.fc-host__editor')).toBeNull();
    key(host, 'Delete');
    key(host, 'Backspace');
    expect(workbook.getValue(target)).toEqual({ kind: 'text', value: 'original' });

    const formulaInput = host.querySelector<HTMLTextAreaElement>('.fc-host__formulabar-input');
    expect(formulaInput).not.toBeNull();
    formulaInput?.focus();
    if (formulaInput) {
      formulaInput.value = 'changed';
      formulaInput.dispatchEvent(new Event('input', { bubbles: true }));
      formulaInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    }
    expect(workbook.getValue(target)).toEqual({ kind: 'text', value: 'original' });

    expect(paste(host, 'pasted').defaultPrevented).toBe(true);
    expect(workbook.getValue(target)).toEqual({ kind: 'text', value: 'original' });

    const canvas = host.querySelector<HTMLElement>('.fc-host__grid canvas');
    canvas?.dispatchEvent(
      new PointerEvent('pointerdown', {
        bubbles: true,
        cancelable: true,
        button: 0,
        clientX: 80,
        clientY: 60,
        pointerId: 1,
      }),
    );
    expect(workbook.getValue(target)).toEqual({ kind: 'text', value: 'original' });
  });

  it('fixed form routes editable input through the controller and rejects mixed paste atomically', async () => {
    sheet = await mountStubSheet({
      policy: fixedFormPolicy([
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
      ]),
      features: { clipboard: true, formulaBar: true, shortcuts: true },
      viewport: {
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 },
        tabNavigation: 'editable',
        tabBoundary: 'leave',
      },
    });
    const { host, instance, workbook } = sheet;
    const editable = addr(0, 0);
    workbook.setNumber(editable, 1);
    workbook.setNumber(addr(0, 1), 99);
    mutators.replaceCells(instance.store, workbook.cells(0));
    mutators.setActive(instance.store, editable);

    key(host, '7');
    const editor = host.querySelector<HTMLTextAreaElement>('.fc-host__editor');
    expect(editor).not.toBeNull();
    if (editor) {
      editor.value = '7';
      editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    }
    expect(workbook.getValue(editable)).toEqual({ kind: 'number', value: 7 });
    expect(instance.history.canUndo()).toBe(true);

    mutators.setActive(instance.store, editable);
    expect(paste(host, '8\t10').defaultPrevented).toBe(true);
    expect(workbook.getValue(editable)).toEqual({ kind: 'number', value: 7 });
    expect(workbook.getValue(addr(0, 1))).toEqual({ kind: 'number', value: 99 });
    expect(instance.history.canUndo()).toBe(true);

    // Tab follows the editable order and stops at the configured component edge.
    mutators.setActive(instance.store, editable);
    const firstTab = key(host, 'Tab');
    expect(firstTab.defaultPrevented).toBe(true);
    expect(instance.store.getState().selection.active).toEqual(addr(0, 2));
    // The next stop is past the final eligible cell, so focus traversal leaves
    // the component and does not move the selection to B1.
    const edgeTab = key(host, 'Tab');
    expect(edgeTab.defaultPrevented).toBe(false);
    expect(instance.store.getState().selection.active).toEqual(addr(0, 2));
  });

  it('selection=false freezes keyboard and pointer navigation while allowing the active cell to edit', async () => {
    sheet = await mountStubSheet({
      policy: { ...fixedFormPolicy([addrRange(0, 0)]), selection: false },
      features: { formulaBar: true, shortcuts: true },
    });
    const { host, instance, workbook } = sheet;
    const active = addr(0, 0);
    mutators.setActive(instance.store, active);
    const before = instance.store.getState().selection;

    for (const keyName of ['ArrowRight', 'Tab', 'Home']) {
      const event = key(host, keyName);
      expect(event.defaultPrevented).toBe(true);
      expect(instance.store.getState().selection).toEqual(before);
    }
    const all = key(host, 'a', { ctrlKey: true });
    expect(all.defaultPrevented).toBe(true);
    expect(instance.store.getState().selection).toEqual(before);

    const canvas = host.querySelector<HTMLElement>('.fc-host__grid canvas');
    const pointer = new PointerEvent('pointerdown', {
      bubbles: true,
      cancelable: true,
      button: 0,
      clientX: 200,
      clientY: 40,
      pointerId: 1,
    });
    canvas?.dispatchEvent(pointer);
    expect(pointer.defaultPrevented).toBe(true);
    expect(instance.store.getState().selection).toEqual(before);

    key(host, '5');
    const editor = host.querySelector<HTMLTextAreaElement>('.fc-host__editor');
    expect(editor).not.toBeNull();
    if (editor) {
      editor.value = '5';
      editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    }
    expect(workbook.getValue(active)).toEqual({ kind: 'number', value: 5 });
    expect(instance.store.getState().selection.active).toEqual(active);
  });

  it('keeps pasted matrix positions when skipIneligible partially applies a fixed form', async () => {
    sheet = await mountStubSheet({
      policy: {
        ...fixedFormPolicy([addrRange(0, 0), addrRange(0, 2)]),
        batchDenied: 'skipIneligible',
      },
      features: { clipboard: true },
    });
    const { host, instance, workbook } = sheet;
    const a1 = addr(0, 0);
    const b1 = addr(0, 1);
    const c1 = addr(0, 2);
    workbook.setNumber(a1, 1);
    workbook.setNumber(b1, 2);
    workbook.setNumber(c1, 3);
    mutators.replaceCells(instance.store, workbook.cells(0));
    mutators.setActive(instance.store, a1);

    expect(paste(host, '8\t9\t10').defaultPrevented).toBe(true);
    expect(workbook.getValue(a1)).toEqual({ kind: 'number', value: 8 });
    expect(workbook.getValue(b1)).toEqual({ kind: 'number', value: 2 });
    expect(workbook.getValue(c1)).toEqual({ kind: 'number', value: 10 });
  });

  it('preflights legacy paste against a viewport range before touching cells', async () => {
    sheet = await mountStubSheet({
      viewport: { range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 } },
      features: { clipboard: true },
    });
    const { host, instance, workbook } = sheet;
    const a1 = addr(0, 0);
    const b1 = addr(0, 1);
    workbook.setNumber(a1, 1);
    workbook.setNumber(b1, 2);
    mutators.replaceCells(instance.store, workbook.cells(0));
    mutators.setActive(instance.store, a1);

    expect(paste(host, '8\t9').defaultPrevented).toBe(true);
    expect(workbook.getValue(a1)).toEqual({ kind: 'number', value: 8 });
    expect(workbook.getValue(b1)).toEqual({ kind: 'number', value: 9 });

    expect(paste(host, '10\t11\t12').defaultPrevented).toBe(true);
    expect(workbook.getValue(a1)).toEqual({ kind: 'number', value: 8 });
    expect(workbook.getValue(b1)).toEqual({ kind: 'number', value: 9 });
  });

  it('uses viewport-only Tab stops and honors a stopping boundary', async () => {
    sheet = await mountStubSheet({
      viewport: {
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
        tabNavigation: 'normal',
        tabBoundary: 'stop',
      },
      features: { shortcuts: true },
    });
    const { host, instance } = sheet;
    mutators.setActive(instance.store, addr(0, 0));

    const first = key(host, 'Tab');
    expect(first.defaultPrevented).toBe(true);
    expect(instance.store.getState().selection.active).toEqual(addr(0, 1));
    const edge = key(host, 'Tab');
    expect(edge.defaultPrevented).toBe(true);
    expect(instance.store.getState().selection.active).toEqual(addr(0, 1));
  });

  it('lets viewport-only Tab traversal leave at a leave boundary', async () => {
    sheet = await mountStubSheet({
      viewport: {
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
        tabNavigation: 'normal',
        tabBoundary: 'leave',
      },
      features: { shortcuts: true },
    });
    const { host, instance } = sheet;
    mutators.setActive(instance.store, addr(0, 1));

    const edge = key(host, 'Tab');
    expect(edge.defaultPrevented).toBe(false);
    expect(instance.store.getState().selection.active).toEqual(addr(0, 1));
  });
});

const addrRange = (row: number, col: number) => ({
  sheet: 0,
  r0: row,
  c0: col,
  r1: row,
  c1: col,
});
