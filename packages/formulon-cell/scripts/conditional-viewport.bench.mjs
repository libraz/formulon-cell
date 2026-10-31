import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright-core';

// Start a demo server first. Timings are local observations, not a CI gate.
const baseURL = process.argv[2] ?? 'http://127.0.0.1:5197';
const rows = Number(process.argv[3] ?? 10000);
if (!Number.isInteger(rows) || rows < 100 || rows > 1048576) {
  throw new Error('Usage: conditional-viewport.bench.mjs [URL] [rule rows: 100..1048576]');
}
const moduleURL = `/@fs${fileURLToPath(new URL('../src/render/conditional.ts', import.meta.url))}`;
const browser = await chromium.launch({ headless: true });
try {
  for (const platform of ['default', 'mac']) {
    for (const kind of ['blanks', 'text-contains', 'date-occurring']) {
      const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
      try {
        const url = new URL(baseURL);
        url.search = new URLSearchParams({ fixture: 'empty', locale: 'en', platform }).toString();
        await page.goto(url.href);
        await page.waitForFunction(
          () =>
            document.querySelector('.fc-host')?.dataset.fcEngineState === 'ready' &&
            window.__fcInst,
        );
        const result = await page.evaluate(
          async ({ kind, rows, moduleURL }) => {
            const { evaluateConditional } = await import(moduleURL);
            const inst = window.__fcInst;
            const frame = () => new Promise(requestAnimationFrame);
            const today = Math.floor(Date.now() / 86400000) + 25569;
            inst.workbook.applyCellPatchAtomic(
              Array.from({ length: 1600 }, (_, index) => ({
                addr: { sheet: 0, row: Math.floor(index / 20), col: index % 20 },
                value:
                  index % 4 === 0
                    ? { kind: 'text', value: 'match' }
                    : { kind: 'number', value: today },
                formula: null,
              })),
            );
            const range = { sheet: 0, r0: 0, c0: 0, r1: rows - 1, c1: 19 };
            const rule = {
              kind,
              range,
              apply: { fill: '#abcdef' },
              ...(kind === 'text-contains' ? { text: 'match' } : {}),
              ...(kind === 'date-occurring' ? { period: 'today' } : {}),
            };
            inst.store.setState((s) => ({
              ...s,
              conditional: { ...s.conditional, rules: [rule] },
            }));
            await frame();
            await frame();
            const state = inst.store.getState();
            const bounds = { r0: 75, c0: 0, r1: 84, c1: 9 };
            const cells = state.data.cells;
            const hadOwnGet = Object.hasOwn(cells, 'get');
            const originalGet = cells.get;
            let reads = 0;
            cells.get = function (key) {
              reads++;
              return originalGet.call(this, key);
            };
            const nativeRaf = window.requestAnimationFrame;
            const callbackDurations = [];
            try {
              const start = performance.now();
              const overlay = evaluateConditional(state, inst.workbook, 'en', bounds);
              const evaluationMs = performance.now() - start;
              const evaluationReads = reads;
              const expectedKey =
                kind === 'blanks' ? '0:80:0' : kind === 'text-contains' ? '0:75:0' : '0:75:1';
              if (overlay.get(expectedKey)?.fill !== '#abcdef') {
                throw new Error(`Conditional output missing at ${expectedKey}`);
              }
              window.requestAnimationFrame = (callback) =>
                nativeRaf((timestamp) => {
                  const start = performance.now();
                  callback(timestamp);
                  callbackDurations.push(performance.now() - start);
                });
              reads = 0;
              const grid = inst.host.querySelector('.fc-host__grid');
              const startRow = inst.store.getState().viewport.rowStart;
              for (let step = 0; step < 6; step++) {
                const s = inst.store.getState();
                grid.dispatchEvent(
                  new WheelEvent('wheel', {
                    deltaY:
                      s.layout.defaultRowHeight * s.viewport.zoom * (s.layout.screenScale ?? 1),
                    bubbles: true,
                    cancelable: true,
                  }),
                );
                await new Promise((resolve) => nativeRaf(resolve));
                await new Promise((resolve) => nativeRaf(resolve));
                if (inst.store.getState().viewport.rowStart !== startRow + step + 1) {
                  throw new Error('Scroll did not advance by one row');
                }
              }
              if (callbackDurations.length === 0) throw new Error('No paint callbacks measured');
              return {
                kind,
                ruleCells: rows * 20,
                evaluationMs,
                evaluationReads,
                overlayEntries: overlay.size,
                visibleCells: 100,
                scrollSteps: 6,
                readsPerScroll: reads / 6,
                paintCallbacks: callbackDurations.length,
                paintCallbackMeanMs:
                  callbackDurations.reduce((sum, ms) => sum + ms, 0) / callbackDurations.length,
              };
            } finally {
              window.requestAnimationFrame = nativeRaf;
              if (hadOwnGet) cells.get = originalGet;
              else delete cells.get;
            }
          },
          { kind, rows, moduleURL },
        );
        process.stdout.write(`${platform}: ${JSON.stringify(result)}\n`);
      } finally {
        await page.close();
      }
    }
  }
} finally {
  await browser.close();
}
