import { beforeAll, describe, expect, it } from 'vitest';
import { loadFormulon } from '../../../src/engine/loader.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import {
  createMacRibbonMenuFactory,
  projectMacFunctionCategoryMenus,
} from '../../../src/toolbar/ribbon/mac/menus.js';

const leafText = (locale: 'de' | 'fr' | 'zh', menuId: string, name: string): string | undefined => {
  const menu = createMacRibbonMenuFactory(locale)(menuId);
  if (!menu) throw new Error(`${menuId} missing for ${locale}`);
  const host = document.createElement('div');
  host.append(menu);
  projectMacFunctionCategoryMenus(host, createSpreadsheetStore(), locale);
  return host
    .querySelector<HTMLButtonElement>(`[data-ribbon-command="mac.function.${name}"]`)
    ?.textContent?.trim();
};

describe('function names in the Mac ribbon menus come from the engine', () => {
  beforeAll(async () => {
    await loadFormulon();
  });

  it('shows the engine names in de and fr with no workbook reader', () => {
    expect(leafText('de', 'mac.formulas.math', 'SUM')).toBe('SUMME');
    expect(leafText('fr', 'mac.formulas.math', 'SUM')).toBe('SOMME');
  });

  it('localizes a function outside the former static table', () => {
    expect(leafText('de', 'mac.formulas.text', 'TEXTJOIN')).toBe('TEXTVERKETTEN');
  });

  it('keeps the canonical name where the engine has no other name', () => {
    expect(leafText('zh', 'mac.formulas.math', 'SUM')).toBe('SUM');
  });
});
