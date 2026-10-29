import { chromium } from 'playwright-core';

const baseURL = process.argv[2] ?? 'http://127.0.0.1:5197';
const selectionOnly = process.argv[3] === 'selection';
const browser = await chromium.launch({ headless: true });
try {
  for (const platform of ['default', 'mac']) {
    for (const axis of ['row', 'col']) {
      const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
      try {
        await page.goto(`${baseURL}/?fixture=empty&locale=en&platform=${platform}`);
        await page.waitForFunction(
          () =>
            document.querySelector('.fc-host')?.dataset.fcEngineState === 'ready' &&
            window.__fcInst,
        );
        const result = await page.evaluate(
          async ({ axis, selectionOnly }) => {
            const inst = window.__fcInst;
            const wb = inst.workbook;
            const host = inst.host;
            const canvas = host.querySelector('.fc-host__canvas');
            const frame = () => new Promise(requestAnimationFrame);
            wb.applyCellPatchAtomic(
              Array.from({ length: 20000 }, (_, index) => ({
                addr: { sheet: 0, row: Math.floor(index / 20), col: index % 20 },
                value: { kind: 'number', value: index + 1 },
                formula: null,
              })),
            );
            host.focus();
            await frame();
            const point = (index) => {
              const { layout, viewport } = inst.store.getState();
              const scale = viewport.zoom * (layout.screenScale ?? 1);
              const bounds = canvas.getBoundingClientRect();
              if (axis === 'col') {
                let x = layout.headerColWidth + layout.outlineRowGutter;
                for (let i = 0; i < index; i++)
                  x += (layout.colWidths.get(i) ?? layout.defaultColWidth) * scale;
                return {
                  clientX:
                    bounds.left +
                    x +
                    ((layout.colWidths.get(index) ?? layout.defaultColWidth) * scale) / 2,
                  clientY: bounds.top + layout.headerRowHeight / 2,
                };
              }
              let y = layout.headerRowHeight + layout.outlineColGutter;
              for (let i = 0; i < index; i++)
                y += (layout.rowHeights.get(i) ?? layout.defaultRowHeight) * scale;
              return {
                clientX: bounds.left + layout.headerColWidth / 2,
                clientY:
                  bounds.top +
                  y +
                  ((layout.rowHeights.get(index) ?? layout.defaultRowHeight) * scale) / 2,
              };
            };
            const select = (index) => {
              const pos = point(index);
              canvas.dispatchEvent(
                new PointerEvent('pointerdown', {
                  ...pos,
                  pointerId: 1,
                  button: 0,
                  buttons: 1,
                  bubbles: true,
                  cancelable: true,
                }),
              );
              canvas.dispatchEvent(
                new PointerEvent('pointerup', {
                  ...pos,
                  pointerId: 1,
                  button: 0,
                  buttons: 0,
                  bubbles: true,
                  cancelable: true,
                }),
              );
              const range = inst.store.getState().selection.range;
              if (
                axis === 'row'
                  ? range.r0 !== index || range.r1 !== index || range.c0 !== 0 || range.c1 !== 16383
                  : range.c0 !== index ||
                    range.c1 !== index ||
                    range.r0 !== 0 ||
                    range.r1 !== 1048575
              ) {
                throw new Error(`Header selection failed: ${axis} ${index}`);
              }
            };
            const closeMenu = () =>
              document.dispatchEvent(
                new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }),
              );
            const openMenu = () => {
              canvas.dispatchEvent(
                new MouseEvent('contextmenu', {
                  ...point(1),
                  button: 2,
                  bubbles: true,
                  cancelable: true,
                }),
              );
              const menu = document.querySelector('.fc-ctxmenu:not(.fc-ctxmenu__sub)');
              if (menu?.style.display !== 'block') throw new Error('Context menu did not open.');
            };
            const calls = {};
            const originals = new Map();
            for (const name of [
              'cells',
              'cellsInRange',
              'getEffectiveStyle',
              'getCellXf',
              'getFontRecord',
              'getFillRecord',
              'getBorderRecord',
            ]) {
              const original = wb[name];
              originals.set(name, original);
              calls[name] = 0;
              wb[name] = function (...args) {
                calls[name] += 1;
                return original.apply(this, args);
              };
            }
            const measure = async (name, count, action) => {
              const durations = [];
              const totals = Object.fromEntries(Object.keys(calls).map((name) => [name, 0]));
              for (let i = 0; i < count; i++) {
                const before = { ...calls };
                const start = performance.now();
                action(i);
                durations.push(performance.now() - start);
                for (const key of Object.keys(calls)) totals[key] += calls[key] - before[key];
                await frame();
              }
              return {
                name,
                count,
                meanMs: durations.reduce((a, b) => a + b, 0) / count,
                firstMs: durations[0],
                lastMs: durations.at(-1),
                callsPerOperation: Object.fromEntries(
                  Object.entries(totals).map(([key, value]) => [key, value / count]),
                ),
              };
            };
            try {
              const selection = await measure('header selection', 20, (i) => select(i % 2));
              if (selectionOnly) return [selection];
              select(0);
              await frame();
              const copy = await measure('whole band copy', 5, () => {
                const transfer = new DataTransfer();
                host.dispatchEvent(
                  new ClipboardEvent('copy', {
                    clipboardData: transfer,
                    bubbles: true,
                    cancelable: true,
                  }),
                );
                if (!transfer.getData('text/plain') || !inst.store.getState().ui.copyRange)
                  throw new Error('Whole band copy failed.');
              });
              select(1);
              await frame();
              const context = await measure('right click with copied band', 5, () => {
                closeMenu();
                openMenu();
              });
              const insert = await measure('insert copied band', 1, () => {
                const button = document.querySelector(
                  '.fc-ctxmenu [data-fc-action="insertCopiedCells"]',
                );
                if (!button || button.disabled)
                  throw new Error('Insert Copied Cells is unavailable.');
                button.click();
                const inserted = wb.getValue({
                  sheet: 0,
                  row: axis === 'row' ? 1 : 0,
                  col: axis === 'col' ? 1 : 0,
                });
                const shifted = wb.getValue({
                  sheet: 0,
                  row: axis === 'row' ? 2 : 0,
                  col: axis === 'col' ? 2 : 0,
                });
                if (
                  inserted.kind !== 'number' ||
                  inserted.value !== 1 ||
                  shifted.kind !== 'number' ||
                  shifted.value !== (axis === 'row' ? 21 : 2)
                )
                  throw new Error(
                    'Copy insertion did not preserve source and shifted target values.',
                  );
              });
              return [selection, copy, context, insert];
            } finally {
              for (const [name, original] of originals) wb[name] = original;
            }
          },
          { axis, selectionOnly },
        );
        process.stdout.write(`${platform}/${axis}: ${JSON.stringify(result)}\n`);
      } finally {
        await page.close();
      }
    }
  }
} finally {
  await browser.close();
}
