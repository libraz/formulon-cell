import { afterEach, describe, expect, it, vi } from 'vitest';
import * as find from '../../../src/commands/find.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachFindReplace } from '../../../src/interact/find-replace.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

vi.mock('../../../src/commands/find.js', async (orig) => {
  const actual = await orig<typeof import('../../../src/commands/find.js')>();
  return { ...actual, findAll: vi.fn(actual.findAll) };
});

describe('find-replace match computation', () => {
  afterEach(() => {
    document.body.innerHTML = '';
    vi.mocked(find.findAll).mockClear();
  });

  it('runs one match scan per Next press', async () => {
    const host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    for (let r = 0; r < 3; r += 1) wb.setText({ sheet: 0, row: r + 1, col: 0 }, 'foo');
    wb.recalc();
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      for (let r = 0; r < 3; r += 1) {
        cells.set(`0:${r + 1}:0`, { value: { kind: 'text', value: 'foo' }, formula: null });
      }
      return { ...s, data: { ...s.data, cells } };
    });
    const handle = attachFindReplace({ getLocale: () => 'en-US', host, store, wb });
    handle.open();
    const input = host.querySelector<HTMLInputElement>('.fc-find input[type="text"]');
    if (!input) throw new Error('no input');
    input.value = 'foo';
    input.dispatchEvent(new Event('input', { bubbles: true }));
    vi.mocked(find.findAll).mockClear();
    host.querySelector<HTMLButtonElement>('.fc-find__btn--next')?.click();
    expect(find.findAll).toHaveBeenCalledTimes(1);
    handle.detach();
  });
});
