import { expect, it } from 'vitest';
import { MAX_COL } from '../../src/engine/address.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { mutators } from '../../src/store/store.js';
import { mountStubSheet } from '../test-utils/mount.js';

function boldXf(workbook: WorkbookHandle): number {
  const normal = workbook.getCellXf(0);
  const font = normal && workbook.getFontRecord(normal.fontIndex);
  if (!normal || !font) throw new Error('Normal font is unavailable');
  return workbook.addXfRecord({
    ...normal,
    fontIndex: workbook.addFontRecord({ ...font, bold: true }),
  });
}

function boldControl(): HTMLInputElement {
  const control = document.querySelector<HTMLInputElement>('input[data-fc-check="bold"]');
  if (!control) throw new Error('Bold control is missing');
  return control;
}

it('includes a distant native row font when the whole worksheet is selected', async () => {
  const workbook = await WorkbookHandle.createDefault({ locale: 'en' });
  const mounted = await mountStubSheet({ workbook, locale: 'en' });
  try {
    expect(workbook.isStub).toBe(false);
    expect(workbook.setRowStyle(0, 999, boldXf(workbook))).toBe(true);
    expect(mounted.instance.store.getState().format.formats.size).toBe(0);
    mutators.selectAll(mounted.instance.store);
    mounted.instance.openFormatDialog('font');
    expect(boldControl().indeterminate).toBe(true);
    expect(mounted.instance.store.getState().format.formats.size).toBe(0);
  } finally {
    mounted.dispose();
    workbook.dispose();
  }
});

it('compares inherited native column fonts across a small selection', async () => {
  const workbook = await WorkbookHandle.createDefault({ locale: 'en' });
  const mounted = await mountStubSheet({ workbook, locale: 'en' });
  try {
    expect(workbook.isStub).toBe(false);
    expect(workbook.setColumnStyle(0, 0, 0, boldXf(workbook))).toBe(true);
    mounted.instance.store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        range: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
      },
    }));
    expect(mounted.instance.store.getState().format.formats.size).toBe(0);
    mounted.instance.openFormatDialog('font');
    expect(boldControl().indeterminate).toBe(true);
    expect(mounted.instance.store.getState().format.formats.size).toBe(0);
  } finally {
    mounted.dispose();
    workbook.dispose();
  }
});

it.each([false, true])(
  'compares actual inherited blanks on a whole sheet (active override=%s)',
  async (overrideActive) => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en' });
    const mounted = await mountStubSheet({ workbook, locale: 'en' });
    try {
      expect(workbook.isStub).toBe(false);
      expect(workbook.setColumnStyle(0, 0, MAX_COL, boldXf(workbook))).toBe(true);
      if (overrideActive) {
        mutators.setCellFormat(
          mounted.instance.store,
          { sheet: 0, row: 0, col: 0 },
          {
            bold: false,
          },
        );
      }
      mutators.selectAll(mounted.instance.store);
      const before = mounted.instance.store.getState().format.formats.size;
      mounted.instance.openFormatDialog('font');
      expect(boldControl().indeterminate).toBe(overrideActive);
      if (!overrideActive) expect(boldControl().checked).toBe(true);
      document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
      expect(mounted.instance.store.getState().format.formats.size).toBe(before);
      const effective = workbook.getEffectiveStyle(0, 100_000, 100);
      const xf = effective && workbook.getCellXf(effective.xfIndex);
      expect(xf && workbook.getFontRecord(xf.fontIndex)?.bold).toBe(true);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  },
);
