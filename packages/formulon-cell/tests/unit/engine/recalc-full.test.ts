import { afterEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { zipEntries, zipStored, zipText } from './zip-entries.js';

const open: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
});

/** The calculation id of the current engine; a book carrying it opens on its cached values. */
const CURRENT_CALC_ID = '191029';

/** A saved workbook whose B1 (`=A1+1`) caches 99 where the formula gives 2, opening without a recalc. */
const staleBook = async (): Promise<Uint8Array> => {
  const source = await WorkbookHandle.createDefault();
  source.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
  source.setFormula({ sheet: 0, row: 0, col: 1 }, '=A1+1');
  source.recalc();
  const entries = zipEntries(source.save());
  source.dispose();
  entries.set(
    'xl/workbook.xml',
    new TextEncoder().encode(
      zipText(entries, 'xl/workbook.xml').replace(
        '</workbook>',
        `<calcPr calcId="${CURRENT_CALC_ID}"/></workbook>`,
      ),
    ),
  );
  const path = 'xl/worksheets/sheet1.xml';
  const xml = zipText(entries, path);
  expect(xml).toContain('<v>2</v>');
  entries.set(path, new TextEncoder().encode(xml.replace('<v>2</v>', '<v>99</v>')));
  return zipStored(entries);
};

const b1 = (wb: WorkbookHandle) => wb.getValue({ sheet: 0, row: 0, col: 1 });

describe('full recalculation', () => {
  it('opens a loaded automatic book on its cached values and keeps them through recalc', async () => {
    const wb = await WorkbookHandle.loadBytes(await staleBook());
    open.push(wb);
    expect(b1(wb)).toMatchObject({ kind: 'number', value: 99 });
    wb.recalc();
    expect(b1(wb)).toMatchObject({ kind: 'number', value: 99 });
  });

  it('recomputes every formula whatever its dirty state', async () => {
    const wb = await WorkbookHandle.loadBytes(await staleBook());
    open.push(wb);
    let recalcs = 0;
    wb.subscribe((event) => {
      if (event.kind === 'recalc') recalcs += 1;
    });

    wb.recalcFull();

    expect(b1(wb)).toMatchObject({ kind: 'number', value: 2 });
    expect(recalcs).toBe(1);
  });
});
