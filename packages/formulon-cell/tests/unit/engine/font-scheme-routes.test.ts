import { FontScheme } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { applyCellStyleToSelection } from '../../../src/commands/cell-styles.js';
import { fillRange } from '../../../src/commands/fill.js';
import { setFont } from '../../../src/commands/format.js';
import { History } from '../../../src/commands/history.js';
import {
  hydrateCellFormatsFromEngine,
  syncCellFormatsToEngine,
} from '../../../src/engine/cell-format-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const at = (col: number) => ({ sheet: 0, row: 0, col });

/** Cell A1 linked to `scheme` (named `name`), B1 on a literal Courier New face. */
async function setup(scheme: FontScheme, name?: string) {
  const wb = await WorkbookHandle.createDefault();
  const base = wb.getFontRecord(0);
  const normal = wb.getCellStyleXf(0);
  if (!base || !normal) throw new Error('engine defaults missing');
  const linked = wb.addFontRecord({ ...base, name: name ?? base.name, scheme, bold: true });
  const literal = wb.addFontRecord({
    ...base,
    name: 'Courier New',
    scheme: FontScheme.None,
    sizePt: 14,
  });
  const xfLinked = wb.addXfRecord({ ...normal, fontIndex: linked, applyFont: true });
  const xfLiteral = wb.addXfRecord({ ...normal, fontIndex: literal, applyFont: true });
  wb.setNumber(at(0), 1);
  wb.setNumber(at(1), 2);
  wb.setCellXfIndex(0, 0, 0, xfLinked);
  wb.setCellXfIndex(0, 0, 1, xfLiteral);
  const store = createSpreadsheetStore();
  hydrateCellFormatsFromEngine(wb, store, 0);
  const fontOf = (col: number) => {
    const xf = wb.getCellXf(wb.getCellXfIndex(0, 0, col) ?? -1);
    return wb.getFontRecord(xf?.fontIndex ?? -1);
  };
  const sync = () => syncCellFormatsToEngine(wb, store, 0, { strict: true });
  return { wb, store, fontOf, sync };
}

describe('theme link carried through store routes', () => {
  it('hydrates the link of a non-body theme font', async () => {
    const { wb, store } = await setup(FontScheme.Major, 'Heading Face');
    try {
      expect(store.getState().format.formats.get('0:0:0')?.fontScheme).toBe('major');
      expect(store.getState().format.formats.get('0:0:1')?.fontScheme).toBeUndefined();
    } finally {
      wb.dispose();
    }
  });

  it('Format Painter-style copy of a major-linked cell onto a literal cell links it', async () => {
    const { wb, store, fontOf, sync } = await setup(FontScheme.Major, 'Heading Face');
    try {
      const src = store.getState().format.formats.get('0:0:0');
      if (!src) throw new Error('no source format');
      store.setState((s) => {
        const formats = new Map(s.format.formats);
        formats.set('0:0:1', { ...src });
        return { ...s, format: { ...s.format, formats } };
      });
      sync();
      expect(fontOf(1)?.scheme).toBe(FontScheme.Major);
      expect(fontOf(1)?.bold).toBe(true);
    } finally {
      wb.dispose();
    }
  });

  it('formats-only fill from a major-linked cell links the target', async () => {
    const { wb, store, fontOf, sync } = await setup(FontScheme.Major, 'Heading Face');
    try {
      fillRange(
        store.getState(),
        wb,
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
        { formatting: 'only', store },
      );
      sync();
      expect(fontOf(1)?.scheme).toBe(FontScheme.Major);
    } finally {
      wb.dispose();
    }
  });

  for (const id of ['normal', 'heading1']) {
    it(`cell style ${id} onto a literal-face cell gives the theme-linked font`, async () => {
      const { wb, store, fontOf, sync } = await setup(FontScheme.Minor);
      try {
        mutators.setActive(store, at(1));
        mutators.setRange(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 });
        expect(applyCellStyleToSelection(store, new History(), id, { getWorkbook: () => wb })).toBe(
          true,
        );
        sync();
        expect(fontOf(1)?.scheme).toBe(FontScheme.Minor);
        expect(fontOf(1)?.name).not.toBe('Courier New');
      } finally {
        wb.dispose();
      }
    });
  }

  it('a user face change on a linked cell makes it literal', async () => {
    const { wb, store, fontOf, sync } = await setup(FontScheme.Major, 'Heading Face');
    try {
      mutators.setActive(store, at(0));
      mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
      setFont(store.getState(), store, { fontFamily: 'Courier New' });
      expect(store.getState().format.formats.get('0:0:0')?.fontScheme).toBeUndefined();
      sync();
      expect(fontOf(0)?.name).toBe('Courier New');
      expect(fontOf(0)?.scheme).toBe(FontScheme.None);
    } finally {
      wb.dispose();
    }
  });
});
