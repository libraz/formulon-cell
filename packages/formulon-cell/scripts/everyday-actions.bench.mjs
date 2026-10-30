import { chromium } from 'playwright-core';

const baseURL = process.argv[2] ?? 'http://127.0.0.1:5197';
const scenarios = ['range paste', 'external paste', 'sparse jump'].filter(
  (scenario) => !process.argv[3] || process.argv[3] === scenario,
);
if (scenarios.length === 0) throw new Error('Unknown benchmark scenario');
const browser = await chromium.launch({ headless: true });
try {
  for (const platform of ['default', 'mac']) {
    for (const scenario of scenarios) {
      const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
      try {
        await page.goto(`${baseURL}/?fixture=empty&locale=en&platform=${platform}`);
        await page.waitForFunction(
          () =>
            document.querySelector('.fc-host')?.dataset.fcEngineState === 'ready' &&
            window.__fcInst,
        );
        const result = await page.evaluate(async (scenario) => {
          const inst = window.__fcInst;
          const wb = inst.workbook;
          const host = inst.host;
          const frame = () => new Promise(requestAnimationFrame);
          wb.applyCellPatchAtomic(
            Array.from({ length: 20000 }, (_, i) => ({
              addr: { sheet: 0, row: Math.floor(i / 20), col: i % 20 },
              value: { kind: 'number', value: i + 1 },
              formula: null,
            })),
          );
          const select = (row, col, r1 = row, c1 = col) => {
            const active = { sheet: 0, row, col };
            inst.store.setState((s) => ({
              ...s,
              selection: {
                ...s.selection,
                active,
                anchor: active,
                range: { sheet: 0, r0: row, c0: col, r1, c1 },
                extraRanges: [],
              },
            }));
          };
          host.focus();
          const durations = [];
          const publications = [];
          for (let trial = 0; trial < 3; trial++) {
            let text;
            if (scenario === 'range paste') {
              select(0, 0, 49, 19);
              const transfer = new DataTransfer();
              host.dispatchEvent(
                new ClipboardEvent('copy', {
                  clipboardData: transfer,
                  bubbles: true,
                  cancelable: true,
                }),
              );
              text = transfer.getData('text/plain');
              if (!text || !inst.store.getState().ui.copyRange) throw new Error('Copy failed');
            } else if (scenario === 'external paste') {
              text = Array.from({ length: 50 }, (_, row) =>
                Array.from({ length: 20 }, (_, col) => String(row * 20 + col + 1)).join('\t'),
              ).join('\n');
            }
            select(
              scenario === 'sparse jump' ? 0 : 60 + trial * 60,
              scenario === 'sparse jump' ? 30 : 0,
            );
            await frame();
            let count = 0;
            const unsub = inst.store.subscribe((s, prev) => {
              if (s.data.cells !== prev.data.cells) count++;
            });
            const start = performance.now();
            if (scenario === 'sparse jump') {
              host.dispatchEvent(
                new KeyboardEvent('keydown', {
                  key: 'ArrowDown',
                  ctrlKey: true,
                  bubbles: true,
                  cancelable: true,
                }),
              );
            } else {
              const transfer = new DataTransfer();
              transfer.setData('text/plain', text);
              host.dispatchEvent(
                new ClipboardEvent('paste', {
                  clipboardData: transfer,
                  bubbles: true,
                  cancelable: true,
                }),
              );
            }
            durations.push(performance.now() - start);
            unsub();
            publications.push(count);
            if (scenario === 'sparse jump') {
              const a = inst.store.getState().selection.active;
              if (a.row !== 1048575 || a.col !== 30) throw new Error('Jump failed');
            } else {
              const value = wb.getValue({ sheet: 0, row: 109 + trial * 60, col: 19 });
              if (value.kind !== 'number' || value.value !== 1000) throw new Error('Paste failed');
            }
            await frame();
          }
          return {
            scenario,
            meanMs: durations.reduce((a, b) => a + b, 0) / durations.length,
            durations,
            publications,
          };
        }, scenario);
        process.stdout.write(`${platform}: ${JSON.stringify(result)}\n`);
      } finally {
        await page.close();
      }
    }
  }
} finally {
  await browser.close();
}
