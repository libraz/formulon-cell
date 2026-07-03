import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it, type Mock, vi } from 'vitest';
import { pasteSpecial } from '../../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../../src/commands/clipboard/snapshot.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { defaultStrings } from '../../../src/i18n/strings.js';
import { attachPasteOptions } from '../../../src/interact/paste-options.js';
import {
  type CellFormat,
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const setRange = (
  store: SpreadsheetStore,
  r0: number,
  c0: number,
  r1: number,
  c1: number,
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      ...s.selection,
      active: { sheet: 0, row: r0, col: c0 },
      anchor: { sheet: 0, row: r0, col: c0 },
      range: { sheet: 0, r0, c0, r1, c1 },
    },
  }));
};

const seedNumber = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  row: number,
  col: number,
  value: number,
  format?: CellFormat,
): void => {
  const addr = { sheet: 0, row, col };
  wb.setNumber(addr, value);
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    const formats = new Map(s.format.formats);
    cells.set(addrKey(addr), { value: { kind: 'number', value }, formula: null });
    if (format) formats.set(addrKey(addr), format);
    return { ...s, data: { ...s.data, cells }, format: { ...s.format, formats } };
  });
};

const formatAt = (store: SpreadsheetStore, row: number, col: number): CellFormat | undefined =>
  store.getState().format.formats.get(addrKey({ sheet: 0, row, col }));

describe('attachPasteOptions', () => {
  let host: HTMLElement;
  let grid: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let onAfterCommit: Mock<() => void>;

  beforeEach(async () => {
    host = document.createElement('div');
    grid = document.createElement('div');
    host.appendChild(grid);
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
    onAfterCommit = vi.fn<() => void>();
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('opens a Paste Options smart button and can switch to Values', () => {
    seedNumber(store, wb, 0, 0, 7, { bold: true, fill: '#fff2cc' });
    seedNumber(store, wb, 2, 2, 2, { italic: true, fill: '#ddebf7' });
    const source = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const before = captureSnapshot(store.getState(), { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 });
    expect(source).not.toBeNull();
    expect(before).not.toBeNull();
    if (!source || !before) throw new Error('missing clipboard snapshots');
    setRange(store, 2, 2, 2, 2);
    pasteSpecial(store.getState(), store, wb, source, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    mutators.replaceCells(store, wb.cells(0));

    const handle = attachPasteOptions({
      host,
      grid,
      store,
      wb,
      strings: defaultStrings,
      onAfterCommit,
    });
    handle.show({
      source,
      before,
      range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
    });

    const button = document.querySelector<HTMLButtonElement>('.fc-paste-options__button');
    expect(button?.style.display).toBe('block');
    button?.click();
    document.querySelector<HTMLButtonElement>('[data-fc-mode="values"]')?.click();

    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 2, col: 2 })).toEqual({ kind: 'number', value: 7 });
    expect(formatAt(store, 2, 2)).toEqual({ italic: true, fill: '#ddebf7' });
    expect(onAfterCommit).toHaveBeenCalledTimes(1);
    handle.detach();
  });

  it('can switch the pasted result to Formatting Only', () => {
    seedNumber(store, wb, 0, 0, 7, { bold: true, fill: '#fff2cc' });
    seedNumber(store, wb, 2, 2, 2, { italic: true, fill: '#ddebf7' });
    const source = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const before = captureSnapshot(store.getState(), { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 });
    expect(source).not.toBeNull();
    expect(before).not.toBeNull();
    if (!source || !before) throw new Error('missing clipboard snapshots');
    setRange(store, 2, 2, 2, 2);
    pasteSpecial(store.getState(), store, wb, source, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    mutators.replaceCells(store, wb.cells(0));

    const handle = attachPasteOptions({
      host,
      grid,
      store,
      wb,
      strings: defaultStrings,
      onAfterCommit,
    });
    handle.show({
      source,
      before,
      range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
    });

    document.querySelector<HTMLButtonElement>('.fc-paste-options__button')?.click();
    document.querySelector<HTMLButtonElement>('[data-fc-mode="formatting"]')?.click();

    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 2, col: 2 })).toEqual({ kind: 'number', value: 2 });
    expect(formatAt(store, 2, 2)).toEqual({ bold: true, fill: '#fff2cc' });
    expect(onAfterCommit).toHaveBeenCalledTimes(1);
    handle.detach();
  });

  it('keeps the smart button and menu within the viewport', () => {
    Object.defineProperty(window, 'innerWidth', { configurable: true, value: 320 });
    Object.defineProperty(window, 'innerHeight', { configurable: true, value: 180 });
    grid.getBoundingClientRect = () =>
      ({
        left: 280,
        top: 150,
        right: 600,
        bottom: 400,
        width: 320,
        height: 250,
        x: 280,
        y: 150,
        toJSON: () => ({}),
      }) as DOMRect;
    seedNumber(store, wb, 0, 0, 7);
    seedNumber(store, wb, 2, 2, 2);
    const source = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const before = captureSnapshot(store.getState(), { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 });
    expect(source).not.toBeNull();
    expect(before).not.toBeNull();
    if (!source || !before) throw new Error('missing clipboard snapshots');

    const handle = attachPasteOptions({
      host,
      grid,
      store,
      wb,
      strings: defaultStrings,
      onAfterCommit,
    });
    const button = document.querySelector<HTMLButtonElement>('.fc-paste-options__button');
    const menu = document.querySelector<HTMLDivElement>('.fc-paste-options__menu');
    expect(button).not.toBeNull();
    expect(menu).not.toBeNull();
    if (!button || !menu) throw new Error('missing paste options elements');
    Object.defineProperty(button, 'offsetWidth', { configurable: true, value: 28 });
    Object.defineProperty(button, 'offsetHeight', { configurable: true, value: 28 });
    Object.defineProperty(menu, 'offsetWidth', { configurable: true, value: 220 });
    Object.defineProperty(menu, 'offsetHeight', { configurable: true, value: 112 });

    handle.show({
      source,
      before,
      range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
    });

    expect(button.style.left).toBe('288px');
    expect(button.style.top).toBe('148px');
    expect(menu.style.left).toBe('96px');
    expect(menu.style.top).toBe('64px');
    handle.detach();
  });

  it('keeps smart button and menu item DOM on the shared floating options helper', () => {
    const source = readFileSync(join(root, 'src/interact/paste-options.ts'), 'utf8');
    const helperSource = readFileSync(join(root, 'src/interact/floating-options-menu.ts'), 'utf8');

    expect(source).toContain('createFloatingOptionsButton({');
    expect(source).toContain('createFloatingOptionsMenuItem({');
    expect(source).not.toContain("const button = document.createElement('button')");
    expect(source).not.toContain("const item = document.createElement('button')");
    expect(helperSource).toContain("import { createInteractionButton } from './chip-button.js'");
    expect(helperSource).not.toContain("document.createElement('button')");
  });

  it('keeps Paste Options close to Excel 365 desktop smart tag geometry', () => {
    const css = readFileSync(join(root, 'src/styles/core/app/overlays/paste-options.css'), 'utf8');

    expect(css).toMatch(
      /\.fc-paste-options__button\s*\{[\s\S]*?width: 22px;[\s\S]*?height: 22px;[\s\S]*?border-radius: 2px;[\s\S]*?0 3px 8px rgba\(0, 0, 0, 0\.16\)/,
    );
    expect(css).toMatch(
      /\.fc-paste-options__menu\s*\{[\s\S]*?min-width: 192px;[\s\S]*?border-radius: 2px;[\s\S]*?font-size: 12px;/,
    );
    expect(css).toMatch(
      /\.fc-paste-options__item\s*\{[\s\S]*?grid-template-columns: 18px minmax\(0, 1fr\);[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 8px;/,
    );
    expect(css).toMatch(
      /\.fc-paste-options__item::before\s*\{[\s\S]*?width: 16px;[\s\S]*?height: 16px;/,
    );
    expect(css).not.toContain('box-shadow: var(--fc-shadow-2)');
  });
});
