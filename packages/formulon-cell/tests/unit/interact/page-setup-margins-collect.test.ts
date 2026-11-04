import { describe, expect, it } from 'vitest';
import { defaultStrings } from '../../../src/i18n/strings.js';
import { createMarginsTab } from '../../../src/interact/page-setup-dialog-tabs/margins.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import { defaultPageSetup } from '../../../src/store/types/page-setup.js';

const build = () => {
  const panel = document.createElement('div');
  document.body.appendChild(panel);
  const tab = createMarginsTab(panel, {
    t: defaultStrings.pageSetup,
    strings: defaultStrings,
    store: createSpreadsheetStore(),
    mac: false,
    on: () => {},
  });
  const input = (label: string): HTMLInputElement => {
    const el = panel.querySelector<HTMLInputElement>(`input[aria-label="${label}"]`);
    if (!el) throw new Error(`missing ${label}`);
    return el;
  };
  return { tab, input };
};

describe('page setup margins tab collect', () => {
  it('clamps negative page margins to zero', () => {
    const { tab, input } = build();
    const t = defaultStrings.pageSetup;
    input(t.marginTop).value = '-1';
    input(t.marginRight).value = '-0.5';
    input(t.marginBottom).value = '-2';
    input(t.marginLeft).value = '-0.1';
    expect(tab.collect().margins).toEqual({ top: 0, right: 0, bottom: 0, left: 0 });
  });

  it('falls back to the default header and footer margins', () => {
    const { tab, input } = build();
    const t = defaultStrings.pageSetup;
    input(t.marginHeader).value = '';
    input(t.marginFooter).value = '';
    const defaults = defaultPageSetup();
    const out = tab.collect();
    expect(out.headerMargin).toBe(defaults.headerMargin);
    expect(out.footerMargin).toBe(defaults.footerMargin);
  });
});
