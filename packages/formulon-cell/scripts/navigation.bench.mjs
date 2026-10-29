import { chromium } from 'playwright-core';

// Start a demo server first, then pass its URL as the first argument.
const baseURL = process.argv[2] ?? 'http://127.0.0.1:5197';
const scenario = process.argv[3] ?? 'move';
const seedCells = Number(process.argv[4] ?? 0);
if (
  !['move', 'scroll', 'edit'].includes(scenario) ||
  !Number.isInteger(seedCells) ||
  seedCells < 0
) {
  throw new Error('Usage: navigation.bench.mjs [URL] [move|scroll|edit] [seed cells]');
}
const browser = await chromium.launch({ headless: true });
try {
  for (const platform of ['default', 'mac']) {
    const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
    try {
      const url = new URL(baseURL);
      url.search = new URLSearchParams({ fixture: 'empty', locale: 'en', platform }).toString();
      await page.goto(url.href);
      await page.waitForFunction(
        () =>
          document.querySelector('.fc-host')?.dataset.fcEngineState === 'ready' && window.__fcInst,
      );
      const result = await page.evaluate(
        async ({ scenario, seedCells }) => {
          const instance = window.__fcInst;
          const host = instance.host;
          if (seedCells > 0) {
            instance.workbook.applyCellPatchAtomic(
              Array.from({ length: seedCells }, (_, index) => ({
                addr: { sheet: 0, row: 1000 + Math.floor(index / 20), col: index % 20 },
                value: { kind: 'number', value: index },
                formula: null,
              })),
            );
            if (instance.store.getState().data.cells.size < seedCells)
              throw new Error('Seed cells were not published.');
          }
          host.focus();
          const nativeRaf = window.requestAnimationFrame.bind(window);
          const frame = () => new Promise((resolve) => nativeRaf(resolve));
          const dispatch = (index) => {
            if (scenario === 'scroll') {
              const state = instance.store.getState();
              host.querySelector('.fc-host__grid').dispatchEvent(
                new WheelEvent('wheel', {
                  deltaY:
                    state.layout.defaultRowHeight *
                    state.viewport.zoom *
                    (state.layout.screenScale ?? 1),
                  bubbles: true,
                  cancelable: true,
                }),
              );
              return;
            }
            if (scenario === 'edit') {
              const address = instance.store.getState().selection.active;
              host.dispatchEvent(
                new KeyboardEvent('keydown', { key: '7', bubbles: true, cancelable: true }),
              );
              const editor = host.querySelector('.fc-host__editor');
              if (!editor) throw new Error('Typing did not open the cell editor.');
              editor.dispatchEvent(
                new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }),
              );
              const committed = instance.store
                .getState()
                .data.cells.get(`${address.sheet}:${address.row}:${address.col}`)?.value;
              if (committed?.kind !== 'number' || committed.value !== 7) {
                throw new Error('Enter did not commit the typed number.');
              }
              return;
            }
            const key = index % 2 === 0 ? 'ArrowDown' : 'ArrowUp';
            host.dispatchEvent(
              new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }),
            );
          };
          for (let i = 0; i < 10; i += 1) {
            dispatch(i);
            await frame();
          }
          const workbook = instance.workbook;
          const calls = {};
          const originals = new Map();
          for (const name of [
            'functionNames',
            'canonicalizeFunctionName',
            'getLambdaText',
            'spillRanges',
            'spillInfo',
            'cells',
            'getEffectiveStyle',
          ]) {
            const original = workbook[name];
            if (typeof original !== 'function') continue;
            originals.set(name, original);
            calls[name] = 0;
            workbook[name] = function (...args) {
              calls[name] += 1;
              return original.apply(this, args);
            };
          }
          const createElement = document.createElement;
          let createdElements = 0;
          document.createElement = function (...args) {
            createdElements += 1;
            return createElement.apply(this, args);
          };
          const raf = window.requestAnimationFrame;
          const paintCallbacks = [];
          window.requestAnimationFrame = (callback) =>
            nativeRaf((timestamp) => {
              const start = performance.now();
              try {
                callback(timestamp);
              } finally {
                paintCallbacks.push(performance.now() - start);
              }
            });
          const synchronous = [];
          const steps = 100;
          const startCell = { ...instance.store.getState().selection.active };
          const startRow = instance.store.getState().viewport.rowStart;
          const navigationStart = performance.now();
          try {
            for (let i = 0; i < steps; i += 1) {
              const start = performance.now();
              dispatch(i);
              synchronous.push(performance.now() - start);
              const active = instance.store.getState().selection.active;
              if (scenario === 'scroll') {
                if (instance.store.getState().viewport.rowStart !== startRow + i + 1) {
                  throw new Error(`Wheel scroll failed at step ${i + 1}.`);
                }
              } else if (
                active.row !==
                  startCell.row + (scenario === 'edit' ? i + 1 : i % 2 === 0 ? 1 : 0) ||
                active.col !== startCell.col
              ) {
                throw new Error(`Arrow navigation failed at step ${i + 1}.`);
              }
              await frame();
            }
            const endCell = instance.store.getState().selection.active;
            if (
              scenario === 'move' &&
              (startCell.row !== endCell.row || startCell.col !== endCell.col)
            ) {
              throw new Error('Alternating navigation did not return to its starting cell.');
            }
            const sorted = [...synchronous].sort((a, b) => a - b);
            const mean = (samples) =>
              samples.reduce((sum, value) => sum + value, 0) / samples.length;
            return {
              steps,
              scenario,
              seedCells,
              navigationDurationMs: performance.now() - navigationStart,
              synchronousMeanMs: mean(synchronous),
              synchronousMedianMs: sorted[Math.floor(sorted.length / 2)],
              synchronousP95Ms: sorted[Math.ceil(sorted.length * 0.95) - 1],
              animationCallbackMeanMs: mean(paintCallbacks),
              animationCallbacks: paintCallbacks.length,
              createdElementsPerStep: createdElements / steps,
              workbookCallsPerStep: Object.fromEntries(
                Object.entries(calls).map(([key, count]) => [key, count / steps]),
              ),
            };
          } finally {
            window.requestAnimationFrame = raf;
            document.createElement = createElement;
            for (const [name, original] of originals) workbook[name] = original;
          }
        },
        { scenario, seedCells },
      );
      process.stdout.write(`${platform}: ${JSON.stringify(result)}\n`);
    } finally {
      await page.close();
    }
  }
} finally {
  await browser.close();
}
