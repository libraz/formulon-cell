import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';

export const newWb = (): Promise<WorkbookHandle> =>
  WorkbookHandle.createDefault({ preferStub: true });

export const cellNumber = (wb: WorkbookHandle, sheet: number, row: number, col: number): number => {
  const v = wb.getValue({ sheet, row, col });
  return v.kind === 'number' ? v.value : Number.NaN;
};

export const cellText = (wb: WorkbookHandle, sheet: number, row: number, col: number): string => {
  const v = wb.getValue({ sheet, row, col });
  return v.kind === 'text' ? v.value : '';
};

export const seedRows = (wb: WorkbookHandle): void => {
  // A1=10, A2=20, A3=30 — three rows on column 0.
  wb.setNumber({ sheet: 0, row: 0, col: 0 }, 10);
  wb.setNumber({ sheet: 0, row: 1, col: 0 }, 20);
  wb.setNumber({ sheet: 0, row: 2, col: 0 }, 30);
  wb.recalc();
};

export const seedCols = (wb: WorkbookHandle): void => {
  // A1=10, B1=20, C1=30 — three cols on row 0.
  wb.setNumber({ sheet: 0, row: 0, col: 0 }, 10);
  wb.setNumber({ sheet: 0, row: 0, col: 1 }, 20);
  wb.setNumber({ sheet: 0, row: 0, col: 2 }, 30);
  wb.recalc();
};
