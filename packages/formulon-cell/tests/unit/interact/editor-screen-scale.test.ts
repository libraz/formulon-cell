import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { InlineEditor } from '../../../src/interact/editor.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

describe('InlineEditor under a platform screen scale', () => {
  let host: HTMLElement;
  let grid: HTMLElement;
  let store: SpreadsheetStore;
  let editor: InlineEditor;

  beforeEach(async () => {
    host = document.createElement('div');
    grid = document.createElement('div');
    host.appendChild(grid);
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    // Mac profile: 100 x 80/3 canonical px cells paint as 75 x 20 screen px;
    // headers are chrome px and stay unscaled.
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        defaultColWidth: 100,
        defaultRowHeight: 80 / 3,
        headerColWidth: 26,
        headerRowHeight: 20,
        screenScale: 0.75,
      },
    }));
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    editor = new InlineEditor({ host, grid, store, wb, onAfterCommit: () => {} });
  });

  afterEach(() => {
    if (editor.isActive()) editor.cancel();
    document.body.replaceChildren();
  });

  const input = (): HTMLTextAreaElement =>
    grid.querySelector('textarea.fc-host__editor') as HTMLTextAreaElement;

  it('frames the cell and paints its 16 px canonical font at 12 screen px', () => {
    const anchor = { sheet: 0, row: 1, col: 1 };
    mutators.setCellFormat(store, anchor, { fontSize: 16 });
    mutators.setActive(store, anchor);
    editor.begin('x');
    // Cell (1, 1) spans x 101..176, y 40..60; the frame sits 1 px outside it.
    expect(input().style.left).toBe('100px');
    expect(input().style.top).toBe('39px');
    expect(input().style.width).toBe('78px');
    expect(input().style.height).toBe('23px');
    expect(input().style.fontSize).toBe('12px');
    expect(input().style.lineHeight).toBe('15px');
  });

  it('grows a two-line edit by two screen rows', () => {
    host.classList.add('fc-host');
    host.dataset.fcPlatform = 'mac';
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    editor.begin('a');
    input().setSelectionRange(1, 1);
    input().dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Enter', altKey: true, cancelable: true, bubbles: true }),
    );
    expect(input().value).toBe('a\n');
    expect(input().style.minHeight).toBe('40px');
  });
});
