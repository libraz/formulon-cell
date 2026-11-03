import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { fillRange } from '../../../src/commands/fill.js';
import { addrKey } from '../../../src/engine/address.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { spreadsheetProfileForLocale } from '../../../src/mount/locale-profile.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

type Group = Record<string, { fwd: string[]; back: string[] | null }>;

// Captured from desktop Excel (ja) by Range.AutoFill of text seeds: `fwd` fills
// down from the top, `back` fills up from the bottom; a `|` separates two seeds.
const fixture = JSON.parse(
  readFileSync(join(__dirname, '../../fixtures/excel-autofill-text.json'), 'utf8'),
) as { lists: Group; numbers: Group };

// Excel types these into a General cell as numbers or dates, so only the
// text-formatted forward fill is a fact.
const FORWARD_ONLY = new Set(['001', '１', '1.', '(1)', '1)', '1-1', '2024-01', '1']);

describe('text autofill matches desktop Excel', () => {
  let wb: WorkbookHandle;
  beforeAll(async () => {
    wb = await WorkbookHandle.createDefault();
    const profile = spreadsheetProfileForLocale('ja-JP', 'mac');
    if (profile) wb.setSpreadsheetProfileId(profile);
  });
  afterAll(() => wb.dispose());

  let column = 0;
  const run = (seeds: string[], rows: number[], dest: [number, number]): string[] => {
    const col = column++;
    const store = createSpreadsheetStore();
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      seeds.forEach((text, i) => {
        const addr = { sheet: 0, row: rows[i] as number, col };
        if (text === '1') {
          wb.setNumber(addr, 1);
          cells.set(addrKey(addr), { value: { kind: 'number', value: 1 }, formula: null });
        } else {
          wb.setText(addr, text);
          cells.set(addrKey(addr), { value: { kind: 'text', value: text }, formula: null });
        }
      });
      return { ...s, data: { ...s.data, cells } };
    });
    const first = rows[0] as number;
    const last = rows[rows.length - 1] as number;
    fillRange(
      store.getState(),
      wb,
      { sheet: 0, r0: first, c0: col, r1: last, c1: col },
      { sheet: 0, r0: dest[0], c0: col, r1: dest[1], c1: col },
      { formatting: 'with', store },
    );
    return Array.from({ length: dest[1] - dest[0] + 1 }, (_, i) => {
      const v = wb.getValue({ sheet: 0, row: dest[0] + i, col });
      return v.kind === 'text' ? v.value : v.kind === 'number' ? String(v.value) : v.kind;
    });
  };

  for (const [group, cases] of Object.entries(fixture)) {
    describe(group, () => {
      for (const [key, want] of Object.entries(cases)) {
        const seeds = key.split('|');
        it(`${key} down`, () => {
          const rows = seeds.map((_, i) => i);
          expect(run(seeds, rows, [0, 5])).toEqual(want.fwd);
        });
        if (want.back && !FORWARD_ONLY.has(key)) {
          it(`${key} up`, () => {
            const rows = seeds.map((_, i) => 10 - seeds.length + i);
            expect(run(seeds, rows, [4, 9])).toEqual(want.back);
          });
        }
      }
    });
  }
});
