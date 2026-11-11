import { describe, expect, it } from 'vitest';
import { commitCellInput } from '../../src/commands/cell-input-commit.js';
import { hyperlinkAt } from '../../src/commands/hyperlinks.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { mountStubSheet } from '../test-utils/mount.js';

const A1 = { sheet: 0, row: 0, col: 0 } as const;
const URL = 'https://example.com/formula-bar-typed';

describe('typed URL from the formula bar', () => {
  it('becomes a hyperlink like one typed in the cell editor', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en' });
    const mounted = await mountStubSheet({ workbook, locale: 'en' });
    try {
      expect(
        commitCellInput({
          store: mounted.instance.store,
          wb: workbook,
          addr: A1,
          raw: URL,
          origin: 'formulaBar',
          typed: true,
        }),
      ).toMatchObject({ status: 'applied' });
      expect(hyperlinkAt(mounted.instance.store.getState(), A1)).toBe(URL);
    } finally {
      mounted.dispose();
    }
  });
});
