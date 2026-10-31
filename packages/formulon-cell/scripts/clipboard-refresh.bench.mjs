import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright-core';

// Start a demo first. This reports local timings without imposing a CI threshold.
const baseURL = process.argv[2] ?? 'http://127.0.0.1:5197';
const seedCells = Number(process.argv[3] ?? 20000);
if (!Number.isInteger(seedCells) || seedCells < 20 || seedCells > 100000 || seedCells % 20) {
  throw new Error('Usage: clipboard-refresh.bench.mjs [URL] [seed cells: multiple of 20]');
}
const moduleURL = `/@fs${fileURLToPath(new URL('../src/interact/clipboard.ts', import.meta.url))}`;
const browser = await chromium.launch({ headless: true });
try {
  for (const platform of ['default', 'mac']) {
    for (const axis of ['column', 'sheet']) {
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
          async ({ axis, seedCells, moduleURL }) => {
            const { attachClipboard } = await import(moduleURL);
            const inst = window.__fcInst;
            const wb = inst.workbook;
            if (!wb.setSpreadsheetProfileId('windows-de_DE')) throw new Error('Profile failed');
            wb.applyCellPatchAtomic(
              Array.from({ length: seedCells }, (_, index) => ({
                addr: { sheet: 0, row: Math.floor(index / 20), col: index % 20 },
                value: { kind: 'blank' },
                formula: '=SUM(1,2)',
              })),
            );
            const range = {
              sheet: 0,
              r0: 0,
              c0: axis === 'column' ? 2 : 0,
              r1: 1048575,
              c1: axis === 'column' ? 2 : 16383,
            };
            inst.store.setState((s) => ({
              ...s,
              selection: { ...s.selection, range, extraRanges: [] },
            }));
            // A second handler on its own host exercises the real clipboard
            // implementation while leaving the demo host's listeners alone.
            const host = document.createElement('div');
            const handle = attachClipboard({ host, store: inst.store, wb, onAfterCommit() {} });
            const localize = wb.localizeFormulaText;
            const cellsInRange = wb.cellsInRange;
            let localizations = 0;
            let rangeQueries = 0;
            wb.localizeFormulaText = function (...args) {
              localizations++;
              return localize.apply(this, args);
            };
            wb.cellsInRange = function (...args) {
              rangeQueries++;
              return cellsInRange.apply(this, args);
            };
            try {
              const transfer = new DataTransfer();
              host.dispatchEvent(
                new ClipboardEvent('copy', {
                  clipboardData: transfer,
                  bubbles: true,
                  cancelable: true,
                }),
              );
              if (!transfer.getData('text/plain').startsWith('=SUMME(1;2)')) {
                throw new Error('External copy was not localized');
              }
              const initialLocalizations = localizations;
              if (initialLocalizations === 0) throw new Error('Localization counter did not bind');
              wb.setFormula({ sheet: 0, row: 0, col: range.c0 }, '=SUM(3,4)');
              const tailRow = seedCells / 20 + 1;
              wb.setFormula({ sheet: 0, row: tailRow, col: 2 }, '=SUM(5,6)');
              const durations = [];
              const localizationCounts = [];
              const queryCounts = [];
              for (let trial = 0; trial < 3; trial++) {
                localizations = 0;
                rangeQueries = 0;
                const start = performance.now();
                const snapshot = handle.getSnapshot();
                durations.push(performance.now() - start);
                localizationCounts.push(localizations);
                queryCounts.push(rangeQueries);
                const first = snapshot?.cells[0]?.[0];
                const last = snapshot?.cells[tailRow]?.[2 - range.c0];
                if (
                  first?.formula !== '=SUM(3,4)' ||
                  first.value.kind !== 'number' ||
                  first.value.value !== 7 ||
                  last?.formula !== '=SUM(5,6)' ||
                  last.value.kind !== 'number' ||
                  last.value.value !== 11
                ) {
                  throw new Error('Refresh lost canonical formulas or latest native values');
                }
                if (snapshot.logicalRange.r1 !== 1048575 || snapshot.rows !== tailRow + 1) {
                  throw new Error('Refresh lost logical or materialized bounds');
                }
              }
              return {
                axis,
                seedCells,
                initialLocalizations,
                durations,
                meanMs: durations.reduce((sum, ms) => sum + ms, 0) / durations.length,
                localizationCounts,
                queryCounts,
              };
            } finally {
              wb.localizeFormulaText = localize;
              wb.cellsInRange = cellsInRange;
              handle.detach();
            }
          },
          { axis, seedCells, moduleURL },
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
