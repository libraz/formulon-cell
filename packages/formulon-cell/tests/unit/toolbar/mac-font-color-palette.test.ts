import { afterEach, describe, expect, it } from 'vitest';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import { createSelectColorRibbon } from '../../../src/toolbar/ribbon/select-color.js';
import { toolbarText } from '../../../src/toolbar/ribbon-model.js';

const open = (command: 'fontColor' | 'fillColor', platform: 'mac' | null) => {
  const host = document.createElement('div');
  if (platform) host.dataset.fcPlatform = platform;
  document.body.appendChild(host);
  const inst = {
    host,
    workbook: { getTheme: () => null },
  } as unknown as SpreadsheetInstance;
  const api = createSelectColorRibbon({
    ribbonLang: 'ja',
    ribbonText: toolbarText('ja') as never,
    pageScaleText: { automatic: '', page: '', pages: '' },
    getInst: () => inst,
    applyRibbonControl: () => undefined,
    currentRibbonControlValue: () => '#000000',
    shouldShowFontOption: () => true,
    createRibbonIcon: () => null,
  });
  const wrap = api.createRibbonColor({
    id: command,
    label: command,
    title: command,
    kind: 'color',
  } as never);
  host.appendChild(wrap);
  wrap.querySelector<HTMLButtonElement>('.fc-tb__rb-color__btn')?.click();
  return (
    document.querySelector('.fc-tb__color-flyout') ?? wrap.querySelector('.fc-tb__color-flyout')
  );
};

// Excel for Mac shows the high-contrast checkbox on the font colour palette too.
describe('Mac colour palettes', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('offers the high-contrast option on the font colour palette on Mac', () => {
    const flyout = open('fontColor', 'mac');
    expect(flyout?.textContent).toContain('ハイ コントラストのみ');
  });

  it('keeps it off the font colour palette elsewhere', () => {
    const flyout = open('fontColor', null);
    expect(flyout?.textContent).not.toContain('ハイ コントラストのみ');
  });
});
