import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { ensureMacInk, getMacInk, type MacInkController } from '../../../src/interact/mac-ink.js';
import { cellRectUnclamped, layoutForView } from '../../../src/render/geometry.js';
import {
  createSpreadsheetStore,
  type SessionIllustration,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const firePointer = (
  target: EventTarget,
  type: string,
  init: PointerEventInit = {},
): PointerEvent => {
  const event = new PointerEvent(type, {
    bubbles: true,
    cancelable: true,
    pointerId: 1,
    button: 0,
    clientX: 48,
    clientY: 30,
    ...init,
  });
  target.dispatchEvent(event);
  return event;
};

const styleRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../../../src/styles/core');
const readStyle = (file: string): string => readFileSync(resolve(styleRoot, file), 'utf8');

/** The z-index tiers of base.css, resolved from `--fc-z-base`. */
const tiers = (): Map<string, number> => {
  const css = readStyle('base.css');
  const base = Number(/--fc-z-base: (\d+);/.exec(css)?.[1]);
  const values = new Map([['base', base]]);
  for (const [, name, offset] of css.matchAll(
    /--fc-z-([a-z]+): calc\(var\(--fc-z-base\) \+ (\d+)\);/g,
  )) {
    values.set(name as string, base + Number(offset));
  }
  return values;
};

/** A z-index value: a number, or `calc(var(--fc-z-tier, fallback) + n)` resolved
 *  against the tiers, whose fallback must agree. */
const resolveZ = (value: string, values: Map<string, number>): number => {
  if (/^\d+$/.test(value.trim())) return Number(value);
  const match = /var\(--fc-z-([a-z]+),\s*(\d+)\)(?:\s*\+\s*(\d+))?/.exec(value);
  if (!match) throw new Error(`unresolved z-index ${value}`);
  const tier = values.get(match[1] as string);
  expect(Number(match[2]), `fallback of --fc-z-${match[1]}`).toBe(tier);
  return (tier ?? Number.NaN) + Number(match[3] ?? 0);
};

/** The literal z-index of `selector` in `file`. */
const cssZ = (file: string, selector: string): number =>
  Number(
    new RegExp(`${selector.replace(/[.]/g, '\\.')} \\{[^}]*?z-index: (\\d+);`).exec(
      readStyle(file),
    )?.[1],
  );

/** The z-index written into an element's style attribute. */
const inlineZ = (element: Element | null): string =>
  /z-index:\s*([^;]+);/.exec(element?.getAttribute('style') ?? '')?.[1] ?? '';

const flushPlatformObserver = async (): Promise<void> => {
  await new Promise<void>((resolve) => setTimeout(resolve, 0));
};

describe('Mac ink controller', () => {
  let host: HTMLElement;
  let grid: HTMLElement;
  let store: SpreadsheetStore;
  let controller: MacInkController | undefined;

  beforeEach(() => {
    host = document.createElement('div');
    host.dataset.fcPlatform = 'mac';
    grid = document.createElement('div');
    grid.className = 'fc-host__grid';
    Object.defineProperty(grid, 'getBoundingClientRect', {
      configurable: true,
      value: () => ({
        left: 0,
        top: 0,
        right: 500,
        bottom: 300,
        width: 500,
        height: 300,
        x: 0,
        y: 0,
        toJSON: () => ({}),
      }),
    });
    host.appendChild(grid);
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    store.setState((state) => ({
      ...state,
      viewport: { ...state.viewport, widthPx: 500 },
    }));
  });

  afterEach(() => {
    controller?.dispose();
    document.body.innerHTML = '';
  });

  it('does not attach to a non-Mac host and reports unsupported Draw tools', () => {
    host.dataset.fcPlatform = 'default';
    expect(ensureMacInk({ host, store })).toBeUndefined();

    host.dataset.fcPlatform = 'mac';
    controller = ensureMacInk({ host, store });
    expect(controller).toBeDefined();
    expect(controller?.supports('lasso')).toBe(false);
    expect(controller?.supports('add-pen')).toBe(false);
    expect(controller?.activate('lasso')).toBe(false);
    expect(controller?.activate('add-pen')).toBe(false);
    expect(host.querySelector('.fc-mac-ink__input')).toBeNull();
  });

  it('stacks the ink over the drawings and under the in-grid panels', () => {
    controller = ensureMacInk({ host, store });
    expect(controller?.activate('pen-black')).toBe(true);
    const values = tiers();
    const chrome = Math.max(
      cssZ('surface/grid-editor.css', '.fc-host__editor'),
      cssZ('surface/scrollbars.css', '.fc-scrolltrack'),
      cssZ('surface/scrollbars.css', '.fc-scrollthumb'),
      cssZ('surface/sheetbar.css', '.fc-host__sheetbar-drop'),
    );
    const drawingCss = /\.fc-drawings \{[^}]*?z-index: ([^;]+);/.exec(
      readStyle('app/overlays/quick-analysis-and-charts.css'),
    )?.[1];
    const drawing = resolveZ(drawingCss ?? '', values);
    const display = resolveZ(inlineZ(host.querySelector('.fc-mac-ink__display')), values);
    const input = resolveZ(inlineZ(host.querySelector('.fc-mac-ink__input')), values);
    const grid = values.get('grid') ?? Number.NaN;
    expect(chrome).toBeGreaterThan(0);
    expect(chrome).toBeLessThan(drawing);
    expect(drawing).toBeLessThan(display);
    expect(display).toBeLessThan(input);
    expect(input).toBeLessThan(grid);
  });

  it('gets an existing controller without creating one', () => {
    const instance = { host, store };
    expect(getMacInk(instance)).toBeUndefined();

    controller = ensureMacInk(instance);
    expect(controller).toBeDefined();
    expect(getMacInk(instance)).toBe(controller);
  });

  it('notifies subscribers when the active tool or trackpad mode changes', () => {
    controller = ensureMacInk({ host, store });
    if (!controller) throw new Error('Mac ink controller did not attach');
    const listener = vi.fn();
    const unsubscribe = controller.subscribe(listener);

    expect(controller.setTool('pen-red')).toBe(true);
    expect(controller.setTrackpadMode(true)).toBe(true);
    controller.deactivate();
    expect(listener).toHaveBeenCalledTimes(3);

    unsubscribe();
    controller.setTool('pencil');
    expect(listener).toHaveBeenCalledTimes(3);
  });

  it('creates a cell-anchored SVG image and makes the stroke undoable', () => {
    controller = ensureMacInk({ host, store });
    expect(controller?.activate('pen-red')).toBe(true);
    const input = host.querySelector<SVGSVGElement>('.fc-mac-ink__input');
    expect(input).toBeTruthy();
    if (!input) throw new Error('missing Mac ink input layer');

    firePointer(input, 'pointerdown', { clientX: 48, clientY: 30 });
    firePointer(input, 'pointermove', { clientX: 84, clientY: 42 });
    firePointer(input, 'pointerup', { clientX: 84, clientY: 42 });

    const item = store.getState().illustrations.illustrations[0] as
      | (SessionIllustration & { macInk?: unknown })
      | undefined;
    expect(item?.id).toMatch(/^mac-ink-0-/);
    expect(item?.src).toMatch(/^data:image\/svg\+xml,/);
    expect(item?.macInk).toMatchObject({ version: 1, tool: 'pen-red' });
    expect(host.querySelector('.fc-mac-ink__display path')).toBeTruthy();

    expect(controller?.deactivate()).toBeUndefined();
    expect(host.querySelector('.fc-mac-ink__input')).toBeNull();
    expect(host.classList.contains('fc-host--mac-ink')).toBe(false);
    expect(controller?.activate('eraser')).toBe(true);
    const stroke = host.querySelector<SVGPathElement>('[data-fc-mac-ink-id]');
    expect(stroke).toBeTruthy();
    if (!stroke) throw new Error('missing rendered ink stroke');
    firePointer(stroke, 'pointerdown', { clientX: 48, clientY: 30 });
    expect(store.getState().illustrations.illustrations).toHaveLength(0);
    expect(controller?.deactivate()).toBeUndefined();

    expect(controller?.activate('pen-black')).toBe(true);
    const secondInput = host.querySelector<SVGSVGElement>('.fc-mac-ink__input');
    if (!secondInput) throw new Error('missing second Mac ink input layer');
    firePointer(secondInput, 'pointerdown', {
      clientX: 48,
      clientY: 30,
    });
    firePointer(secondInput, 'pointerup', {
      clientX: 48,
      clientY: 30,
    });
    expect(store.getState().illustrations.illustrations).toHaveLength(1);
    expect((controller as MacInkController).isActive()).toBe(true);
    expect((controller as MacInkController).getTool()).toBe('pen-black');
  });

  it('reprojects strokes when the viewport scrolls and zooms', () => {
    controller = ensureMacInk({ host, store });
    controller?.activate('pencil');
    const input = host.querySelector<SVGSVGElement>('.fc-mac-ink__input');
    if (!input) throw new Error('missing Mac ink input layer');
    firePointer(input, 'pointerdown', { clientX: 48, clientY: 30 });
    firePointer(input, 'pointermove', { clientX: 84, clientY: 42 });
    firePointer(input, 'pointerup', { clientX: 84, clientY: 42 });

    const path = (): string => host.querySelector('[data-fc-mac-ink-id]')?.getAttribute('d') ?? '';
    const atTop = path();
    store.setState((state) => ({
      ...state,
      viewport: { ...state.viewport, rowStart: 1 },
    }));
    const afterScroll = path();
    expect(afterScroll).not.toBe(atTop);
    store.setState((state) => ({
      ...state,
      viewport: { ...state.viewport, zoom: 2 },
    }));
    expect(path()).not.toBe(afterScroll);
  });

  it('clips the input and rendered paths to the data viewport', () => {
    controller = ensureMacInk({ host, store });
    controller?.activate('pen-black');
    const input = host.querySelector<SVGSVGElement>('.fc-mac-ink__input');
    const display = host.querySelector<SVGSVGElement>('.fc-mac-ink__display');
    if (!input || !display) throw new Error('missing Mac ink layers');

    const inputClip = input.querySelector('clipPath rect');
    const displayClip = display.querySelector('clipPath rect');
    if (!inputClip || !displayClip) throw new Error('missing Mac ink clip paths');
    expect(Number(inputClip.getAttribute('x'))).toBeGreaterThan(0);
    expect(Number(inputClip.getAttribute('y'))).toBeGreaterThan(0);
    expect(Number(inputClip.getAttribute('width'))).toBeLessThan(500);
    expect(Number(inputClip.getAttribute('height'))).toBeLessThan(300);
    expect(input.querySelector('g[clip-path]')).toBeTruthy();
    expect(display.querySelector('g[clip-path]')).toBeTruthy();

    store.setState((state) => ({
      ...state,
      ui: { ...state.ui, showHeaders: false },
    }));
    const noHeadersClip = input.querySelector('clipPath rect');
    if (!noHeadersClip) throw new Error('missing no-header clip path');
    expect(Number(noHeadersClip.getAttribute('x'))).toBe(0);
    expect(Number(noHeadersClip.getAttribute('y'))).toBe(0);
  });

  it('clips strokes to their anchored freeze pane, including RTL layout', () => {
    store.setState((state) => ({
      ...state,
      layout: { ...state.layout, freezeRows: 1, freezeCols: 1 },
      viewport: { ...state.viewport, rowStart: 1, colStart: 1, widthPx: 500 },
    }));
    controller = ensureMacInk({ host, store });
    controller?.activate('pen-black');
    const input = host.querySelector<SVGSVGElement>('.fc-mac-ink__input');
    if (!input) throw new Error('missing Mac ink input layer');

    const drawInCell = (row: number, col: number): void => {
      const state = store.getState();
      const rect = cellRectUnclamped(layoutForView(state), state.viewport, row, col);
      const x = rect.x + Math.max(1, rect.w / 2);
      const y = rect.y + Math.max(1, rect.h / 2);
      firePointer(input, 'pointerdown', { clientX: x, clientY: y });
      firePointer(input, 'pointerup', { clientX: x, clientY: y });
    };

    drawInCell(1, 1);
    drawInCell(0, 1);
    drawInCell(1, 0);
    drawInCell(0, 0);
    const paths = Array.from(host.querySelectorAll<SVGPathElement>('[data-fc-mac-ink-id]'));
    expect(paths).toHaveLength(4);
    expect(paths[0]?.parentElement?.getAttribute('clip-path')).toMatch(/-body\)$/);
    expect(paths[1]?.parentElement?.getAttribute('clip-path')).toMatch(/-frozen-rows\)$/);
    expect(paths[2]?.parentElement?.getAttribute('clip-path')).toMatch(/-frozen-cols\)$/);
    expect(paths[3]?.parentElement?.getAttribute('clip-path')).toMatch(/-corner\)$/);

    store.setState((state) => ({
      ...state,
      ui: { ...state.ui, rightToLeft: true },
    }));
    const display = host.querySelector<SVGSVGElement>('.fc-mac-ink__display');
    if (!display) throw new Error('missing Mac ink display layer');
    const rtlBody = display.querySelector<SVGRectElement>('clipPath[id$="-body"] rect');
    const rtlFrozenCols = display.querySelector<SVGRectElement>(
      'clipPath[id$="-frozen-cols"] rect',
    );
    if (!rtlBody || !rtlFrozenCols) throw new Error('missing RTL pane clips');
    expect(Number(rtlBody.getAttribute('x'))).toBe(0);
    expect(Number(rtlFrozenCols.getAttribute('x'))).toBeGreaterThan(0);
  });

  it('cancels an in-progress stroke and transitions cleanly with the host platform', async () => {
    controller = ensureMacInk({ host, store });
    expect(controller?.activate('pen-red')).toBe(true);
    const input = host.querySelector<SVGSVGElement>('.fc-mac-ink__input');
    if (!input) throw new Error('missing Mac ink input layer');

    firePointer(input, 'pointerdown');
    firePointer(input, 'pointermove', { clientX: 84, clientY: 42 });
    firePointer(input, 'pointercancel');
    expect(store.getState().illustrations.illustrations).toHaveLength(0);

    firePointer(input, 'pointerdown');
    firePointer(input, 'pointerup');
    expect(store.getState().illustrations.illustrations).toHaveLength(1);
    const display = host.querySelector<SVGSVGElement>('.fc-mac-ink__display');
    if (!display) throw new Error('missing Mac ink display layer');
    expect(display.hasAttribute('hidden')).toBe(false);

    host.dataset.fcPlatform = 'default';
    await flushPlatformObserver();
    expect(controller?.isActive()).toBe(false);
    expect(host.querySelector('.fc-mac-ink__input')).toBeNull();
    expect(host.classList.contains('fc-host--mac-ink')).toBe(false);
    // The viewport-aware projection stays visible while the generic session
    // image remains suppressed by mac-ink.css on every platform.
    expect(display.hasAttribute('hidden')).toBe(false);
    expect(store.getState().illustrations.illustrations).toHaveLength(1);

    host.dataset.fcPlatform = 'mac';
    await flushPlatformObserver();
    expect(display.hasAttribute('hidden')).toBe(false);
    expect(host.querySelector('.fc-mac-ink__input')).toBeNull();
    expect(controller?.getTool()).toBeNull();
    expect(host.querySelector('[data-fc-mac-ink-id]')).toBeTruthy();
    expect(controller?.activate()).toBe(true);
    expect(host.querySelector<SVGSVGElement>('.fc-mac-ink__input')?.dataset.fcMacInkTool).toBe(
      'pen-red',
    );
  });

  it('keeps the preview and committed stroke widths aligned at zoom', () => {
    store.setState((state) => ({
      ...state,
      viewport: { ...state.viewport, zoom: 2 },
    }));
    controller = ensureMacInk({ host, store });
    controller?.activate('pen-black');
    const input = host.querySelector<SVGSVGElement>('.fc-mac-ink__input');
    if (!input) throw new Error('missing Mac ink input layer');
    firePointer(input, 'pointerdown');
    const preview = input.querySelector('path');
    expect(preview?.getAttribute('stroke-width')).toBe('4.5');
    firePointer(input, 'pointerup');
    expect(host.querySelector('[data-fc-mac-ink-id]')?.getAttribute('stroke-width')).toBe('4.5');
  });

  it('does not mutate a protected or restricted sheet', async () => {
    const protectedController = ensureMacInk({ host, store });
    if (!protectedController) throw new Error('Mac ink controller did not attach');
    controller = protectedController;
    protectedController.activate('pen-black');
    setProtectedSheet(store, 0, true);
    const input = host.querySelector<SVGSVGElement>('.fc-mac-ink__input');
    if (!input) throw new Error('missing Mac ink input layer');
    firePointer(input, 'pointerdown');
    firePointer(input, 'pointerup');
    expect(store.getState().illustrations.illustrations).toHaveLength(0);

    protectedController.dispose();
    setProtectedSheet(store, 0, false);
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const commands = new InteractionController({ store, getWb: () => wb, history: new History() });
    const unregister = registerInteractionController(store, commands);
    commands.setPolicy({ defaultOperation: 'allow', operations: { object: false } });
    const restricted = ensureMacInk({ host, store });
    if (!restricted) throw new Error('Mac ink controller did not reattach');
    controller = restricted;
    restricted.activate('pen-black');
    const restrictedInput = host.querySelector<SVGSVGElement>('.fc-mac-ink__input');
    if (!restrictedInput) throw new Error('missing restricted Mac ink input layer');
    firePointer(restrictedInput, 'pointerdown');
    firePointer(restrictedInput, 'pointerup');
    expect(store.getState().illustrations.illustrations).toHaveLength(0);
    unregister();
    commands.dispose();
    wb.dispose();
  });

  it('restores a stroke through shared history after erasing', () => {
    const history = new History();
    const attached = ensureMacInk({ host, store, history });
    if (!attached) throw new Error('Mac ink controller did not attach');
    controller = attached;
    attached.activate('pen-black');
    const input = host.querySelector<SVGSVGElement>('.fc-mac-ink__input');
    if (!input) throw new Error('missing Mac ink input layer');
    firePointer(input, 'pointerdown');
    firePointer(input, 'pointerup');
    expect(store.getState().illustrations.illustrations).toHaveLength(1);

    attached.activate('eraser');
    const stroke = host.querySelector<SVGPathElement>('[data-fc-mac-ink-id]');
    if (!stroke) throw new Error('missing rendered ink stroke');
    firePointer(stroke, 'pointerdown');
    expect(store.getState().illustrations.illustrations).toHaveLength(0);
    expect(history.undo()).toBe(true);
    expect(store.getState().illustrations.illustrations).toHaveLength(1);
    expect(attached.getTool()).toBe('eraser');
  });

  describe('saved with the workbook', () => {
    const PNG_BASE64 =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';

    beforeEach(() => {
      vi.stubGlobal('Path2D', class {});
      vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(
        () =>
          ({
            scale: vi.fn(),
            stroke: vi.fn(),
          }) as unknown as CanvasRenderingContext2D,
      );
      vi.spyOn(HTMLCanvasElement.prototype, 'toDataURL').mockReturnValue(
        `data:image/png;base64,${PNG_BASE64}`,
      );
    });

    afterEach(() => {
      vi.unstubAllGlobals();
      vi.restoreAllMocks();
    });

    it('stores a stroke as a PNG picture that undo and the eraser remove', async () => {
      const workbook = await WorkbookHandle.createDefault();
      try {
        const history = new History();
        workbook.attachHistory(history);
        controller = ensureMacInk({ host, store, history, workbook });
        controller?.activate('pen-red');
        const input = host.querySelector<SVGSVGElement>('.fc-mac-ink__input');
        if (!input) throw new Error('missing Mac ink input layer');
        firePointer(input, 'pointerdown', { clientX: 48, clientY: 30 });
        firePointer(input, 'pointermove', { clientX: 84, clientY: 42 });
        firePointer(input, 'pointerup', { clientX: 84, clientY: 42 });

        const pictures = workbook.listDrawingObjects(0).filter((object) => object.kind === 0);
        expect(pictures).toHaveLength(1);
        const item = store.getState().illustrations.illustrations[0] as
          | (SessionIllustration & { macInk?: unknown; nativeObjectId?: number })
          | undefined;
        expect(item?.nativeObjectId).toBe(pictures[0]?.objectId);
        expect(item?.macInk).toMatchObject({ tool: 'pen-red' });
        expect(workbook.getImage(0, item?.nativeObjectId ?? 0)?.bytes.length).toBeGreaterThan(0);

        const restored = await WorkbookHandle.loadBytes(workbook.save());
        try {
          expect(restored.listDrawingObjects(0).filter((object) => object.kind === 0)).toHaveLength(
            1,
          );
        } finally {
          restored.dispose();
        }

        history.undo();
        expect(workbook.listDrawingObjects(0).filter((object) => object.kind === 0)).toHaveLength(
          0,
        );
        expect(store.getState().illustrations.illustrations).toHaveLength(0);
        history.redo();
        const redone = store.getState().illustrations.illustrations[0] as
          | (SessionIllustration & { macInk?: unknown })
          | undefined;
        expect(redone?.macInk).toMatchObject({ tool: 'pen-red' });

        controller?.deactivate();
        controller?.activate('eraser');
        const stroke = host.querySelector<SVGPathElement>('[data-fc-mac-ink-id]');
        if (!stroke) throw new Error('missing rendered ink stroke');
        firePointer(stroke, 'pointerdown', { clientX: 48, clientY: 30 });
        expect(workbook.listDrawingObjects(0).filter((object) => object.kind === 0)).toHaveLength(
          0,
        );
      } finally {
        workbook.dispose();
      }
    });
  });
});
