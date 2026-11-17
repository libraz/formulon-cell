import { FontScheme } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import {
  hydrateCellFormatsFromEngine,
  syncCellFormatsToEngine,
} from '../../../src/engine/cell-format-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { type CellFormat, createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const ADDR = { sheet: 0, row: 0, col: 0 };

type Variant = { readonly label: string; readonly name: string; readonly scheme: FontScheme };
// A minor-scheme font on the body face, and a major-scheme font on another face.
const VARIANTS: readonly Variant[] = [
  { label: 'body face', name: '', scheme: FontScheme.Minor },
  { label: 'heading face', name: 'Heading Face', scheme: FontScheme.Major },
];

async function withLinkedCell(
  variant: Variant,
  run: (ctx: {
    wb: WorkbookHandle;
    apply: (patch: Partial<CellFormat>) => void;
    scheme: () => FontScheme | undefined;
    name: () => string | undefined;
  }) => void,
): Promise<void> {
  const wb = await WorkbookHandle.createDefault();
  try {
    const base = wb.getFontRecord(0);
    const normal = wb.getCellStyleXf(0);
    if (!base || !normal) throw new Error('engine defaults missing');
    const fontIndex = wb.addFontRecord({
      ...base,
      name: variant.name || base.name,
      scheme: variant.scheme,
    });
    const xf = wb.addXfRecord({ ...normal, fontIndex, applyFont: true });
    wb.setNumber(ADDR, 1);
    expect(wb.setCellXfIndex(0, 0, 0, xf)).toBe(true);
    const store = createSpreadsheetStore();
    hydrateCellFormatsFromEngine(wb, store, 0);
    const fontOf = () =>
      wb.getFontRecord(wb.getCellXf(wb.getCellXfIndex(0, 0, 0) ?? -1)?.fontIndex ?? -1);
    run({
      wb,
      apply: (patch) => {
        mutators.setCellFormat(store, ADDR, patch);
        syncCellFormatsToEngine(wb, store, 0, { strict: true });
      },
      scheme: () => fontOf()?.scheme,
      name: () => fontOf()?.name,
    });
  } finally {
    wb.dispose();
  }
}

describe('theme-linked cell font through the store sync', () => {
  for (const variant of VARIANTS) {
    describe(variant.label, () => {
      const cases: ReadonlyArray<readonly [string, Partial<CellFormat>]> = [
        ['size step', { fontSize: 16 }],
        ['bold', { bold: true }],
        ['italic', { italic: true }],
        ['underline', { underline: 'single' }],
        ['strike', { strike: true }],
        ['font colour', { color: '#ff0000' }],
        ['whole-font copy (same face)', { bold: true, fontSize: 18, color: '#00aa00' }],
      ];
      for (const [label, patch] of cases) {
        it(`keeps the scheme on ${label}`, async () => {
          await withLinkedCell(variant, ({ apply, scheme, name }) => {
            const before = name();
            apply(patch);
            expect(scheme()).toBe(variant.scheme);
            expect(name()).toBe(before);
          });
        });
      }
    });
  }

  it('drops the link when the user picks another face', async () => {
    await withLinkedCell(VARIANTS[0] as Variant, ({ apply, scheme, name }) => {
      apply({ fontFamily: 'Courier New' });
      expect(name()).toBe('Courier New');
      expect(scheme()).toBe(FontScheme.None);
    });
  });
});
