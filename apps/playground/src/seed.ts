import type { WorkbookHandle } from '@libraz/formulon-cell';

export const seedWorkbook = (wb: WorkbookHandle): void => {
  wb.setText({ sheet: 0, row: 0, col: 0 }, 'item');
  wb.setText({ sheet: 0, row: 0, col: 1 }, 'qty');
  wb.setText({ sheet: 0, row: 0, col: 2 }, 'unit');
  wb.setText({ sheet: 0, row: 0, col: 3 }, 'subtotal');
  wb.setText({ sheet: 0, row: 0, col: 4 }, 'tax (8%)');
  const rows = [
    ['paper', 24, 0.42],
    ['vermillion ink', 6, 12.5],
    ['rule pen', 2, 8.9],
    ['draftsman pad', 1, 24.0],
    ['eraser', 3, 1.25],
  ] as const;
  rows.forEach(([name, qty, unit], i) => {
    const r = i + 1;
    wb.setText({ sheet: 0, row: r, col: 0 }, name);
    wb.setNumber({ sheet: 0, row: r, col: 1 }, qty);
    wb.setNumber({ sheet: 0, row: r, col: 2 }, unit);
    wb.setFormula({ sheet: 0, row: r, col: 3 }, `=B${r + 1}*C${r + 1}`);
    wb.setFormula({ sheet: 0, row: r, col: 4 }, `=D${r + 1}*0.08`);
  });
  wb.setText({ sheet: 0, row: 7, col: 0 }, 'total');
  wb.setFormula({ sheet: 0, row: 7, col: 3 }, '=SUM(D2:D6)');
  wb.setFormula({ sheet: 0, row: 7, col: 4 }, '=SUM(E2:E6)');
  wb.setFormula({ sheet: 0, row: 8, col: 3 }, '=D8+E8');
  wb.setText({ sheet: 0, row: 8, col: 0 }, 'with tax');
  wb.recalc();
};
