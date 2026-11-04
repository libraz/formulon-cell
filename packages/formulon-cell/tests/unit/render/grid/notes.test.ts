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

  it('sets the Mac note card: 105 x 77 pt, 1 pt black frame and leader, bold text, no author line', () => {
    // Excel (Mac, 2x): the card spans device x 49-258 / y 13-166 beside a
    // cell whose top-right corner is (22, 36); the leader runs from that
    // corner to the card's top-left; "note text" is bold black at x 55.
    const store = createSpreadsheetStore();
    store.setState((s) => ({ ...s, layout: { ...s.layout, screenScale: 0.75 } }));
    setWidth(store, 1000);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 2, col: 1 },
      { comment: 'note text', commentVisible: true },
    );
    const state = store.getState();
    const view = { ...state, layout: layoutForView(state) };
    const cols = buildColLayout(view.layout, view.viewport);
    const rows = buildRowLayout(view.layout, view.viewport);
    const cell = cellRectIn(view.layout, cols, rows, 2, 1);
    const fills: { style: string; rect: number[] }[] = [];
    const strokes: { style: string; width: number; rect?: number[]; path?: number[][] }[] = [];
    const texts: {
      text: string;
      font: string;
      style: string;
      x: number;
      y: number;
      base: string;
    }[] = [];
    let path: number[][] = [];
    const ctx = {
      fillStyle: '',
      strokeStyle: '',
      lineWidth: 1,
      font: '',
      textBaseline: 'alphabetic',
      textAlign: 'left',
      save(): void {},
      restore(): void {},
      beginPath(): void {
        path = [];
      },
      rect(): void {},
      clip(): void {},
      moveTo(x: number, y: number): void {
        path.push([x, y]);
      },
      lineTo(x: number, y: number): void {
        path.push([x, y]);
      },
      stroke(this: { strokeStyle: string; lineWidth: number }): void {
        strokes.push({ style: this.strokeStyle, width: this.lineWidth, path });
      },
      strokeRect(this: { strokeStyle: string; lineWidth: number }, ...r: number[]): void {
        strokes.push({ style: this.strokeStyle, width: this.lineWidth, rect: r });
      },
      measureText: (text: string) => ({ width: text.length * 6 }),
      fillRect(this: { fillStyle: string }, ...r: number[]): void {
        fills.push({ style: this.fillStyle, rect: r });
      },
      fillText(
        this: { font: string; fillStyle: string; textBaseline: string },
        text: string,
        x: number,
        y: number,
      ): void {
        texts.push({ text, font: this.font, style: this.fillStyle, x, y, base: this.textBaseline });
      },
    } as unknown as CanvasRenderingContext2D;

    paintVisibleNotes(ctx, view, cols, rows);

    const corner = cell.x + cell.w;
    const box = fills.find((f) => f.style === NOTE_BOX_FILL)?.rect ?? [];
    expect(box).toEqual([corner + 13.5, cell.y - 11.5, 105, 77]);
    const frame = strokes.find((s) => s.rect);
    expect(frame).toMatchObject({ style: '#000000', width: 1 });
    expect(frame?.rect).toEqual([corner + 14, cell.y - 11, 104, 76]);
    const leader = strokes.find((s) => s.path && s.path.length === 2);
    expect(leader).toMatchObject({ style: '#000000', width: 1 });
    expect(leader?.path).toEqual([
      [corner, cell.y],
      [corner + 13.5, cell.y - 11.5],
    ]);
    expect(texts).toHaveLength(1);
    expect(texts[0]).toMatchObject({ text: 'note text', style: '#000000', base: 'alphabetic' });
    expect(texts[0]?.font).toMatch(/^bold 9px /);
    expect(texts[0]?.x).toBeCloseTo(corner + 13.5 + 2.5, 5);
    expect(texts[0]?.y).toBeCloseTo(cell.y - 11.5 + 15.5, 5);
  });
});
