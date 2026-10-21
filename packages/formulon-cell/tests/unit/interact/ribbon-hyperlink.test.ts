import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { setHyperlink } from '../../../src/commands/hyperlinks.js';
import { en } from '../../../src/i18n/strings.js';
import { executeRibbonHyperlinkAction } from '../../../src/interact/ribbon-hyperlink.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const workbook = {} as Parameters<typeof executeRibbonHyperlinkAction>[0]['workbook'];

const runOpen = (url: string) => {
  const store = createSpreadsheetStore();
  setHyperlink(store, { sheet: 0, row: 0, col: 0 }, url);
  return executeRibbonHyperlinkAction({
    store,
    workbook,
    history: new History(),
    action: 'open',
    strings: {
      linkOpen: en.ribbonMenu.linkOpen,
      linkNoHyperlink: en.ribbonMenu.linkNoHyperlink,
    },
  });
};

describe('executeRibbonHyperlinkAction', () => {
  it('returns a trimmed URL for an allowed target', () => {
    expect(runOpen('  HTTPS://example.test/path  ')).toEqual({
      kind: 'open-url',
      url: 'HTTPS://example.test/path',
    });
  });

  it('opens in-document targets through the same route', () => {
    expect(runOpen('#Sheet2!A1')).toEqual({ kind: 'open-url', url: '#Sheet2!A1' });
  });

  it('refuses unsafe targets through the ribbon route', () => {
    const result = runOpen('javascript:alert(1)');
    expect(result.kind).toBe('report');
    expect(result).not.toEqual(expect.objectContaining({ kind: 'open-url' }));
  });
});
