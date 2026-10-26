import { describe, expect, it } from 'vitest';
import {
  buildColLayout,
  buildRowLayout,
  cellRectIn,
  gridOriginX,
  gridOriginY,
  layoutForView,
} from '../../../../src/render/geometry.js';
import {
  NOTE_BOX_FILL,
  paintVisibleNotes,
  visibleNotes,
} from '../../../../src/render/grid/notes.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';

function makeCtxSpy(): {
  ctx: CanvasRenderingContext2D;
  fills: { style: string; rect: number[] }[];
  texts: string[];
  clips: number[][];
} {
  const fills: { style: string; rect: number[] }[] = [];
  const texts: string[] = [];
  const clips: number[][] = [];
  let lastRect: number[] = [];
  let fillStyle = '';
  const ctx = {
    get fillStyle(): string {
      return fillStyle;
    },
    set fillStyle(v: string) {
      fillStyle = v;
    },
    strokeStyle: '',
    lineWidth: 1,
    font: '',
    textBaseline: 'alphabetic',
    textAlign: 'left',
    save(): void {},
    restore(): void {},
    beginPath(): void {},
    rect(x: number, y: number, w: number, h: number): void {
      lastRect = [x, y, w, h];
    },
    clip(): void {
      clips.push(lastRect);
    },
    moveTo(): void {},
    lineTo(): void {},
    stroke(): void {},
    strokeRect(): void {},
    measureText: (text: string) => ({ width: text.length * 6 }),
    fillRect(x: number, y: number, w: number, h: number): void {
      fills.push({ style: fillStyle, rect: [x, y, w, h] });
    },
    fillText(text: string): void {
      texts.push(text);
    },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, fills, texts, clips };
}

const paint = (
  store: ReturnType<typeof createSpreadsheetStore>,
  at: { row: number; col: number } = { row: 2, col: 1 },
) => {
  const state = store.getState();
  const view = { ...state, layout: layoutForView(state) };
  const cols = buildColLayout(view.layout, view.viewport);
  const rows = buildRowLayout(view.layout, view.viewport);
  const spy = makeCtxSpy();
  paintVisibleNotes(spy.ctx, view, cols, rows);
  return {
    ...spy,
    layout: view.layout,
    cols,
    rows,
    cell: cellRectIn(view.layout, cols, rows, at.row, at.col),
  };
};

const showNote = (
  store: ReturnType<typeof createSpreadsheetStore>,
  row: number,
  col: number,
): void => {
  mutators.setCellFormat(store, { sheet: 0, row, col }, { comment: 'n', commentVisible: true });
};

const setWidth = (store: ReturnType<typeof createSpreadsheetStore>, widthPx: number): void => {
  store.setState((s) => ({ ...s, viewport: { ...s.viewport, widthPx } }));
};

describe('paintVisibleNotes', () => {
  it('draws a shown note beside its cell and skips hidden ones', () => {
    const store = createSpreadsheetStore();
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 2, col: 1 },
      {
        comment: 'Shown text',
        commentAuthor: 'Alice',
        commentVisible: true,
      },
    );
    mutators.setCellFormat(store, { sheet: 0, row: 5, col: 4 }, { comment: 'Hidden text' });
    const { fills, texts, cell } = paint(store);
    const boxes = fills.filter((f) => f.style === NOTE_BOX_FILL);
    expect(boxes).toHaveLength(1);
    expect(boxes[0]?.rect[0]).toBeGreaterThan(cell.x + cell.w);
    expect(texts.join(' ')).toContain('Alice:');
    expect(texts.join(' ')).toContain('Shown text');
    expect(texts.join(' ')).not.toContain('Hidden text');
  });

  it('paints nothing when no note is shown', () => {
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { comment: 'note' });
    expect(paint(store).fills).toHaveLength(0);
  });

  it('indexes shown notes once per formats map', () => {
    const store = createSpreadsheetStore();
    showNote(store, 2, 1);
    mutators.setCellFormat(store, { sheet: 0, row: 4, col: 4 }, { bold: true });
    const formats = store.getState().format.formats;
    const first = visibleNotes(formats);
    expect(first.map((n) => n.addr)).toEqual([{ sheet: 0, row: 2, col: 1 }]);
    expect(visibleNotes(formats)).toBe(first);
    showNote(store, 6, 6);
    expect(visibleNotes(store.getState().format.formats)).toHaveLength(2);
  });

  it('clips a note in the frozen rows to that pane, below the column headers', () => {
    const store = createSpreadsheetStore();
    setWidth(store, 1000);
    mutators.setFreezePanes(store, 3, 0);
    showNote(store, 0, 1);
    const { clips, layout, rows } = paint(store, { row: 0, col: 1 });
    const pane = clips[0] ?? [];
    expect(pane[1]).toBe(gridOriginY(layout));
    expect(pane[3]).toBe(rows.frozenTotal);
  });

  it('clips a note in the body to the body pane', () => {
    const store = createSpreadsheetStore();
    setWidth(store, 1000);
    mutators.setFreezePanes(store, 3, 2);
    showNote(store, 5, 4);
    const { clips, layout, rows, cols } = paint(store, { row: 5, col: 4 });
    const pane = clips[0] ?? [];
    expect(pane[0]).toBe(gridOriginX(layout) + cols.frozenTotal);
    expect(pane[1]).toBe(gridOriginY(layout) + rows.frozenTotal);
  });

  it('places the note on the left and keeps it off the row headers on a right-to-left sheet', () => {
    const store = createSpreadsheetStore();
    setWidth(store, 1000);
    mutators.setRightToLeft(store, true);
    showNote(store, 2, 1);
    const { fills, clips, layout, cell } = paint(store);
    const box = fills.find((f) => f.style === NOTE_BOX_FILL)?.rect ?? [];
    expect((box[0] ?? 0) + (box[2] ?? 0)).toBeLessThan(cell.x);
    const pane = clips[0] ?? [];
    expect((pane[0] ?? 0) + (pane[2] ?? 0)).toBe(1000 - gridOriginX(layout));
  });
});
