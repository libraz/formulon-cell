import { describe, expect, it } from 'vitest';
import type { CellValue } from '../../../src/engine/types.js';
import { formatCell } from '../../../src/engine/value.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachWatchPanel } from '../../../src/interact/watch-panel.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const VALUE: CellValue = { kind: 'number', value: 1234567.891 };

const valueCellText = (profile: string | null, hostLocale?: string): string => {
  const wb = {
    spreadsheetProfileId: () => profile,
    getValue: () => VALUE,
    sheetName: () => 'Sheet1',
    definedNames: () => [],
    cellFormula: () => null,
  } as unknown as WorkbookHandle;
  const store = createSpreadsheetStore();
  mutators.addWatch(store, { sheet: 0, row: 0, col: 0 });
  const host = document.createElement('div');
  document.body.appendChild(host);
  const panel = attachWatchPanel({
    host,
    store,
    getWb: () => wb,
    getLocale: hostLocale ? () => hostLocale : undefined,
  });
  panel.refresh();
  const text = host.querySelector('.fc-watch__value')?.textContent ?? '';
  panel.detach();
  host.remove();
  return text;
};

describe('watch panel display locale without a native renderer', () => {
  const cases: [string, string | null, string | undefined, string][] = [
    ['ja_JP', 'default-ja_JP', 'fr-FR', 'ja-JP'],
    ['de_DE', 'mac-de_DE', 'fr-FR', 'de-DE'],
    ['fr_FR', 'default-fr_FR', 'de-DE', 'fr-FR'],
    ['zh_CN', 'default-zh_CN', 'de-DE', 'zh-CN'],
    ['zh_TW', 'default-zh_TW', 'de-DE', 'zh-TW'],
    ['ko_KR', 'default-ko_KR', 'de-DE', 'ko-KR'],
    ['th_TH', 'default-th_TH', 'de-DE', 'th-TH'],
    ['en_US defers to the host locale', 'default-en_US', 'de-DE', 'de-DE'],
    ['en_US without a host locale', 'default-en_US', undefined, 'en-US'],
    ['no profile uses the host locale', null, 'fr-FR', 'fr-FR'],
    ['unknown profile uses the host locale', 'default-xx_XX', 'fr-FR', 'fr-FR'],
  ];
  for (const [name, profile, host, expected] of cases) {
    it(name, () => {
      expect(valueCellText(profile, host)).toBe(formatCell(VALUE, expected));
    });
  }
});

it('distinguishes locales (guards the characterization against a vacuous formatter)', () => {
  expect(formatCell(VALUE, 'de-DE')).not.toBe(formatCell(VALUE, 'en-US'));
});
