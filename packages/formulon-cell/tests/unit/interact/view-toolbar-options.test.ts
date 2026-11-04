import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { attachViewToolbar } from '../../../src/interact/view-toolbar.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('attachViewToolbar sheet-view options', () => {
  it('keeps the option nodes across unrelated store changes and rebuilds on view changes', () => {
    const toolbar = document.createElement('div');
    document.body.appendChild(toolbar);
    const store = createSpreadsheetStore();
    const handle = attachViewToolbar({
      toolbar,
      store,
      wb: { capabilities: {} } as unknown as WorkbookHandle,
      history: new History(),
      strings: en,
    });
    const select = toolbar.querySelector<HTMLSelectElement>('select[aria-label="Views"]');
    if (!select) throw new Error('missing sheet-view select');
    const before = Array.from(select.options);
    expect(before.length).toBeGreaterThan(0);

    mutators.setActive(store, { sheet: 0, row: 3, col: 2 });
    mutators.setZoom(store, 1.5);
    expect(Array.from(select.options)).toEqual(before);
    expect(Array.from(select.options).every((o, i) => o === before[i])).toBe(true);

    toolbar.querySelector<HTMLButtonElement>('button[aria-label="Save"]')?.click();
    expect(select.options.length).toBe(before.length + 1);
    handle.detach();
    toolbar.remove();
  });
});
