import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  attachValidationAlert,
  attachValidationList,
  attachValidationPrompt,
} from '../../../src/interact/validation.js';
import { setValidationChevron } from '../../../src/render/grid/hit-state.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const popover = (): HTMLElement | null =>
  document.querySelector<HTMLElement>('.fc-validation-list');
const items = (): HTMLElement[] =>
  Array.from(document.querySelectorAll<HTMLElement>('.fc-validation-list__item'));
const prompt = (): HTMLElement | null =>
  document.querySelector<HTMLElement>('.fc-validation-prompt');

const stubGridRect = (grid: HTMLElement): void => {
  grid.getBoundingClientRect = (): DOMRect =>
    ({ left: 0, top: 0, right: 200, bottom: 200, width: 200, height: 200, x: 0, y: 0 }) as DOMRect;
};

const firePointerInChevron = (
  grid: HTMLElement,
  chevronX: number,
  chevronY: number,
): PointerEvent => {
  const e = new PointerEvent('pointerdown', {
    clientX: chevronX + 2,
    clientY: chevronY + 2,
    button: 0,
    bubbles: true,
    cancelable: true,
    pointerId: 1,
  });
  grid.dispatchEvent(e);
  return e;
};

describe('attachValidationList', () => {
  let grid: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let onAfterCommit: () => void;

  beforeEach(async () => {
    grid = document.createElement('div');
    document.body.appendChild(grid);
    stubGridRect(grid);
    store = createSpreadsheetStore();
    wb = await newWb();
    onAfterCommit = vi.fn<() => void>();
  });

  afterEach(() => {
    setValidationChevron(null);
    while (document.body.firstChild) document.body.removeChild(document.body.firstChild);
  });

  it('chevron pointerdown on a list-validated cell opens a dropdown of allowed values', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 1, col: 1 },
      { validation: { kind: 'list', source: ['alpha', 'beta', 'gamma'] } },
    );
    setValidationChevron({ rect: { x: 50, y: 30, w: 12, h: 12 }, row: 1, col: 1 });

    const handle = attachValidationList({ grid, store, wb, onAfterCommit });
    expect(popover()).toBeNull();
    firePointerInChevron(grid, 50, 30);
    expect(popover()).not.toBeNull();
    expect(items().map((i) => i.textContent)).toEqual(['alpha', 'beta', 'gamma']);
    expect(popover()?.getAttribute('role')).toBe('listbox');
    expect(document.activeElement).toBe(items()[0]);
    expect(items()[0]?.getAttribute('aria-selected')).toBe('true');
    handle.detach();
  });

  it('clicking an item writes the chosen value to the workbook and notifies onAfterCommit', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { validation: { kind: 'list', source: ['x', 'y'] } },
    );
    setValidationChevron({ rect: { x: 70, y: 50, w: 12, h: 12 }, row: 0, col: 0 });

    const handle = attachValidationList({ grid, store, wb, onAfterCommit });
    firePointerInChevron(grid, 70, 50);
    const item = items().find((i) => i.textContent === 'y');
    item?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
    expect(onAfterCommit).toHaveBeenCalledTimes(1);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'y' });
    // Popover closes after a pick.
    expect(popover()).toBeNull();
    handle.detach();
  });

  it('does not write a validation-list pick into a locked protected cell', () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { validation: { kind: 'list', source: ['x', 'y'] } },
    );
    setProtectedSheet(store, 0, true);
    setValidationChevron({ rect: { x: 70, y: 50, w: 12, h: 12 }, row: 0, col: 0 });

    try {
      const handle = attachValidationList({ grid, store, wb, onAfterCommit });
      firePointerInChevron(grid, 70, 50);
      const item = items().find((i) => i.textContent === 'y');
      item?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));

      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
      expect(onAfterCommit).not.toHaveBeenCalled();
      expect(warn).toHaveBeenCalledTimes(1);
      expect(popover()).toBeNull();
      handle.detach();
    } finally {
      warn.mockRestore();
    }
  });

  it('Escape on the document closes an open popover', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { validation: { kind: 'list', source: ['only'] } },
    );
    setValidationChevron({ rect: { x: 10, y: 10, w: 12, h: 12 }, row: 0, col: 0 });

    const handle = attachValidationList({ grid, store, wb, onAfterCommit });
    firePointerInChevron(grid, 10, 10);
    expect(popover()).not.toBeNull();
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(popover()).toBeNull();
    handle.detach();
  });

  it('supports arrow-key selection and Enter commit', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { validation: { kind: 'list', source: ['small', 'medium', 'large'] } },
    );
    setValidationChevron({ rect: { x: 10, y: 10, w: 12, h: 12 }, row: 0, col: 0 });

    const handle = attachValidationList({ grid, store, wb, onAfterCommit });
    firePointerInChevron(grid, 10, 10);
    expect(document.activeElement).toBe(items()[0]);

    items()[0]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(document.activeElement).toBe(items()[1]);
    expect(items()[1]?.getAttribute('aria-selected')).toBe('true');

    items()[1]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
    expect(document.activeElement).toBe(items()[2]);

    items()[2]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(onAfterCommit).toHaveBeenCalledTimes(1);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'large',
    });
    expect(popover()).toBeNull();
    handle.detach();
  });

  it('opens with the current cell value selected when it is in the list', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { validation: { kind: 'list', source: ['small', 'medium', 'large'] } },
    );
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'text', value: 'medium' });
    setValidationChevron({ rect: { x: 10, y: 10, w: 12, h: 12 }, row: 0, col: 0 });

    const handle = attachValidationList({ grid, store, wb, onAfterCommit });
    firePointerInChevron(grid, 10, 10);

    expect(document.activeElement).toBe(items()[1]);
    expect(items()[1]?.getAttribute('aria-selected')).toBe('true');
    expect(items()[0]?.getAttribute('aria-selected')).toBe('false');
    handle.detach();
  });

  it('Escape restores focus to the grid opener', () => {
    grid.tabIndex = 0;
    grid.focus();
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { validation: { kind: 'list', source: ['only'] } },
    );
    setValidationChevron({ rect: { x: 10, y: 10, w: 12, h: 12 }, row: 0, col: 0 });

    const handle = attachValidationList({ grid, store, wb, onAfterCommit });
    firePointerInChevron(grid, 10, 10);
    expect(popover()).not.toBeNull();
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));

    expect(popover()).toBeNull();
    expect(document.activeElement).toBe(grid);
    handle.detach();
  });

  it('mousedown outside the popover closes it', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { validation: { kind: 'list', source: ['only'] } },
    );
    setValidationChevron({ rect: { x: 10, y: 10, w: 12, h: 12 }, row: 0, col: 0 });

    const handle = attachValidationList({ grid, store, wb, onAfterCommit });
    firePointerInChevron(grid, 10, 10);
    expect(popover()).not.toBeNull();

    const outside = document.createElement('div');
    document.body.appendChild(outside);
    outside.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
    expect(popover()).toBeNull();
    handle.detach();
  });

  it('detach removes the pointerdown listener so subsequent chevron taps do nothing', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { validation: { kind: 'list', source: ['only'] } },
    );
    setValidationChevron({ rect: { x: 10, y: 10, w: 12, h: 12 }, row: 0, col: 0 });

    const handle = attachValidationList({ grid, store, wb, onAfterCommit });
    handle.detach();
    firePointerInChevron(grid, 10, 10);
    expect(popover()).toBeNull();
  });

  it('pointerdown outside the chevron rect does not open the popover', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { validation: { kind: 'list', source: ['x'] } },
    );
    setValidationChevron({ rect: { x: 100, y: 100, w: 12, h: 12 }, row: 0, col: 0 });

    const handle = attachValidationList({ grid, store, wb, onAfterCommit });
    // Far-away click.
    firePointerInChevron(grid, 0, 0);
    expect(popover()).toBeNull();
    handle.detach();
  });

  it('keeps the validation list close to Excel 365 desktop dropdown geometry', () => {
    const css = readFileSync(
      join(root, 'src/styles/core/app/popups/validation-and-chooser.css'),
      'utf8',
    );

    expect(css).toMatch(
      /\.fc-validation-list\s*\{[\s\S]*?border-radius: 0;[\s\S]*?box-shadow:[\s\S]*?font-size: 12px;/,
    );
    expect(css).toMatch(
      /\.fc-validation-list__item\s*\{[\s\S]*?min-height: 22px;[\s\S]*?padding: 3px 8px;/,
    );
    expect(css).toMatch(
      /\.fc-validation-list__item:hover,[\s\S]*?\.fc-validation-list__item\[aria-selected="true"\]\s*\{[\s\S]*?background: var\(--fc-bg-hover/,
    );
    expect(css).not.toContain(
      '.fc-validation-list__item[aria-selected="true"] {\n    background: var(--fc-accent-soft',
    );
  });
});

describe('attachValidationPrompt', () => {
  let grid: HTMLElement;
  let store: SpreadsheetStore;

  beforeEach(() => {
    grid = document.createElement('div');
    document.body.appendChild(grid);
    stubGridRect(grid);
    store = createSpreadsheetStore();
  });

  afterEach(() => {
    while (document.body.firstChild) document.body.removeChild(document.body.firstChild);
  });

  it('shows the active cell validation input message', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        validation: {
          kind: 'list',
          source: ['Open', 'Closed'],
          promptTitle: 'Status',
          promptMessage: 'Choose a workflow state.',
        },
      },
    );

    const handle = attachValidationPrompt({ grid, store });

    expect(prompt()).not.toBeNull();
    expect(prompt()?.hidden).toBe(false);
    expect(prompt()?.textContent).toContain('Status');
    expect(prompt()?.textContent).toContain('Choose a workflow state.');
    handle.detach();
  });

  it('hides when showInputMessage is false or the selection moves away', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        validation: {
          kind: 'whole',
          op: 'between',
          a: 1,
          b: 10,
          promptMessage: 'Enter 1 to 10.',
          showInputMessage: false,
        },
      },
    );
    const handle = attachValidationPrompt({ grid, store });
    expect(prompt()).toBeNull();

    mutators.setCellFormat(
      store,
      { sheet: 0, row: 1, col: 0 },
      {
        validation: {
          kind: 'whole',
          op: 'between',
          a: 1,
          b: 10,
          promptMessage: 'Enter 1 to 10.',
        },
      },
    );
    mutators.setActive(store, { sheet: 0, row: 1, col: 0 });
    expect(prompt()?.hidden).toBe(false);

    mutators.setActive(store, { sheet: 0, row: 2, col: 0 });
    expect(prompt()?.hidden).toBe(true);
    handle.detach();
  });

  it('keeps the validation input prompt on a restrained desktop popover surface', () => {
    const css = readFileSync(
      join(root, 'src/styles/core/app/popups/validation-and-chooser.css'),
      'utf8',
    );

    expect(css).toMatch(
      /\.fc-validation-prompt\s*\{[\s\S]*?padding: 7px 9px;[\s\S]*?border-radius: 1px;[\s\S]*?box-shadow:/,
    );
    expect(css).toMatch(
      /\.fc-validation-prompt__title\s*\{[\s\S]*?font-weight: 600;[\s\S]*?margin-bottom: 4px;/,
    );
  });
});

describe('attachValidationAlert', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    while (document.body.firstChild) document.body.removeChild(document.body.firstChild);
  });

  it('shows a validation error alert with a custom title and message', () => {
    const handle = attachValidationAlert({
      host,
      labels: { ok: 'OK', stop: 'Stop', warning: 'Warning', information: 'Information' },
    });

    handle.show({
      severity: 'stop',
      title: 'Invalid status',
      message: 'Use Open, Closed, or Hold.',
    });

    const dialog = document.querySelector<HTMLElement>('.fc-valdlg');
    expect(dialog).not.toBeNull();
    expect(dialog?.hidden).toBe(false);
    expect(dialog?.getAttribute('aria-label')).toBe('Invalid status');
    expect(dialog?.querySelector('.fc-valdlg__panel')).not.toBeNull();
    expect(dialog?.textContent).toContain('Use Open, Closed, or Hold.');
    handle.detach();
  });

  it('keeps the validation alert on a compact desktop dialog surface', () => {
    const css = readFileSync(
      join(root, 'src/styles/core/app/popups/validation-and-chooser.css'),
      'utf8',
    );

    expect(css).toMatch(
      /\.fc-valdlg__panel\s*\{[\s\S]*?width: min\(360px, calc\(100vw - 48px\)\);[\s\S]*?border-radius: 2px;/,
    );
    expect(css).toMatch(
      /\.fc-valdlg__panel \.fc-fmtdlg__header\s*\{[\s\S]*?min-height: 30px;[\s\S]*?padding: 7px 12px 6px;[\s\S]*?font-size: 12px;/,
    );
    expect(css).toMatch(
      /\.fc-valdlg__panel \.app__dlg__message\s*\{[\s\S]*?margin: 0;[\s\S]*?font-size: 12px;[\s\S]*?line-height: 1.35;/,
    );
    expect(css).toMatch(
      /\.fc-valdlg__panel \.fc-fmtdlg__footer\s*\{[\s\S]*?min-height: 38px;[\s\S]*?padding: 6px 12px 10px;/,
    );
  });
});
