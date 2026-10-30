import { chromium } from 'playwright-core';

const baseURL = process.argv[2] ?? 'http://127.0.0.1:5197';
const seedCells = Number(process.argv[3] ?? 20000);
const requireCopy = process.argv[4] === '--require-copy';
if (!Number.isInteger(seedCells) || seedCells < 1) throw new Error('Invalid cell count');
const scenarios = [
  { axis: 'sheet', action: 'copy' },
  { axis: 'sheet', action: 'bold' },
  { axis: 'sheet', action: 'fontGrow' },
  { axis: 'sheet', action: 'comma' },
  { axis: 'row', action: 'bold' },
  { axis: 'column', action: 'bold' },
];
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
        const result = await page.evaluate(
          async ({ axis, action, seedCells, requireCopy }) => {
            const inst = window.__fcInst;
            const wb = inst.workbook;
            const host = inst.host;
            const canvas = host.querySelector('.fc-host__canvas');
            wb.applyCellPatchAtomic(
              Array.from({ length: seedCells }, (_, index) => ({
                addr: { sheet: 0, row: Math.floor(index / 20), col: index % 20 },
                value: { kind: 'number', value: index + 1 },
                formula: null,
              })),
            );
            await new Promise(requestAnimationFrame);
            host.focus();
            const state = inst.store.getState();
            const bounds = canvas.getBoundingClientRect();
            const scale = state.viewport.zoom * (state.layout.screenScale ?? 1);
            const pos = {
              clientX:
                bounds.left +
                (axis === 'column'
                  ? state.layout.headerColWidth +
                    state.layout.outlineRowGutter +
                    (state.layout.defaultColWidth * scale) / 2
                  : state.layout.headerColWidth / 2),
              clientY:
                bounds.top +
                (axis === 'row'
                  ? state.layout.headerRowHeight +
                    state.layout.outlineColGutter +
                    (state.layout.defaultRowHeight * scale) / 2
                  : state.layout.headerRowHeight / 2),
            };
            for (const type of ['pointerdown', 'pointerup'])
              canvas.dispatchEvent(
                new PointerEvent(type, {
                  ...pos,
                  pointerId: 1,
                  button: 0,
                  buttons: type === 'pointerdown' ? 1 : 0,
                  bubbles: true,
                  cancelable: true,
                }),
              );
            const selected = inst.store.getState().selection.range;
            if (
              selected.r0 !== 0 ||
              selected.c0 !== 0 ||
              selected.r1 !== (axis === 'row' ? 0 : 1048575) ||
              selected.c1 !== (axis === 'column' ? 0 : 16383)
            )
              throw new Error('Selection failed');
            await new Promise(requestAnimationFrame);
            const counts = {};
            const originals = new Map();
            for (const name of [
              'cells',
              'physicalCells',
              'getCellXfIndex',
              'getEffectiveStyle',
              'setCellXfIndex',
            ]) {
              const original = wb[name];
              originals.set(name, original);
              counts[name] = 0;
              wb[name] = function (...args) {
                counts[name]++;
                return original.apply(this, args);
              };
            }
            let copied = null;
            const start = performance.now();
            try {
              if (action === 'copy') {
                const transfer = new DataTransfer();
                host.dispatchEvent(
                  new ClipboardEvent('copy', {
                    clipboardData: transfer,
                    bubbles: true,
                    cancelable: true,
                  }),
                );
                copied = transfer.getData('text/plain');
              } else {
                const button = document.querySelector(`[data-ribbon-command="${action}"]`);
                if (!button || button.disabled) throw new Error(`Missing command: ${action}`);
                button.click();
              }
              const durationMs = performance.now() - start;
              const measuredCounts = { ...counts };
              if (action === 'copy') {
                if (requireCopy && !copied) throw new Error('Whole-sheet copy failed');
              } else {
                const farRow = axis === 'row' ? 0 : 1048575;
                const farCol = axis === 'column' ? 0 : 16383;
                const effective = wb.getEffectiveStyle(0, farRow, farCol);
                const xf = effective && wb.getCellXf(effective.xfIndex);
                const font = xf && wb.getFontRecord(xf.fontIndex);
                if (action === 'bold' && !font?.bold) throw new Error('Bold was not applied');
                if (action === 'fontGrow' && !(font?.size > 11))
                  throw new Error('Font size was not applied');
                if (action === 'comma' && effective?.numFmtCode !== '#,##0.00')
                  throw new Error('Number format was not applied');
              }
              const after = inst.store.getState();
              return {
                axis,
                action,
                seedCells,
                durationMs,
                counts: measuredCounts,
                copied: copied === null ? null : copied.length > 0,
                copiedChars: copied?.length ?? null,
                formats: after.format.formats.size,
                sheetFontOverride: after.format.sheetFontOverrides?.has(0) ?? false,
                sheetAxisOverride: after.format.sheetAxisOverrides?.has(0) ?? false,
              };
            } finally {
              for (const [name, original] of originals) wb[name] = original;
            }
          },
          { ...scenario, seedCells, requireCopy },
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
