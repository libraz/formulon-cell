import { afterEach, expect, it } from 'vitest';
import { MAX_COL, MAX_ROW } from '../../../src/engine/address.js';
import { attachFormatDialog } from '../../../src/interact/format-dialog.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

afterEach(() => {
  document.body.innerHTML = '';
});

it.each([
  ['worksheet', MAX_ROW, MAX_COL],
  ['column', MAX_ROW, 0],
  ['row', 0, MAX_COL],
] as const)('shows mixed font for one bold cell and implicit blanks in a whole %s', (_, r1, c1) => {
  const store = createSpreadsheetStore();
  mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
  mutators.selectAll(store);
  store.setState((state) => ({
    ...state,
    selection: {
      ...state.selection,
      range: { sheet: 0, r0: 0, c0: 0, r1, c1 },
    },
  }));
  const host = document.createElement('div');
  document.body.appendChild(host);
  const dialog = attachFormatDialog({ host, store });
  try {
    dialog.open('font');
    const bold = document.querySelector<HTMLInputElement>('input[data-fc-check="bold"]');
    expect(bold).not.toBeNull();
    expect(bold?.indeterminate).toBe(true);
    expect(store.getState().format.formats.size).toBe(1);
  } finally {
    dialog.detach();
  }
});
