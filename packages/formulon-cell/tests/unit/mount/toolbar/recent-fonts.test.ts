import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createDefaultRibbonHelpers } from '../../../../src/mount/toolbar-defaults.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { clearRecentFonts } from '../../../../src/toolbar/ribbon/font-availability.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { seedText } from './fixtures.js';

describe('Recently Used Fonts in the Font list', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    clearRecentFonts();
    sheet = await mountStubSheet({
      workbook: await WorkbookHandle.createDefault({ locale: 'en-US' }),
      locale: 'en',
    });
    host = document.createElement('div');
    document.body.appendChild(host);
  });
  afterEach(() => {
    sheet.dispose();
    host.remove();
    clearRecentFonts();
  });

  const openList = (): HTMLElement => {
    const existing = document.querySelector<HTMLElement>('.fc-tb__rb-dd__list');
    if (existing) return existing;
    host.querySelector<HTMLButtonElement>('[data-ribbon-command="fontFamily"] button')?.click();
    const list = document.querySelector<HTMLElement>('.fc-tb__rb-dd__list');
    if (!list) throw new Error('font list did not open');
    return list;
  };
  const headings = (list: HTMLElement): string[] =>
    Array.from(list.querySelectorAll('.fc-tb__rb-dd__section'), (el) => el.textContent ?? '');

  it('adds a Recently Used section once a font has been picked, newest first', () => {
    seedText(sheet, 0, 0, 'x');
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: createDefaultRibbonHelpers(sheet.instance),
    });
    expect(headings(openList())).not.toContain('Recently Used Fonts');
    document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    for (const family of ['Arial', 'Times New Roman']) {
      const list = openList();
      list.querySelector<HTMLButtonElement>(`.fc-tb__rb-dd__opt[data-value="${family}"]`)?.click();
    }
    const list = openList();
    const names = headings(list);
    expect(names.indexOf('Recently Used Fonts')).toBeGreaterThan(-1);
    const recent = Array.from(
      list.querySelectorAll('.fc-tb__rb-dd__opt[data-recent="true"]'),
      (el) => el.getAttribute('data-value'),
    );
    expect(recent).toEqual(['Times New Roman', 'Arial']);
    tb.dispose();
  });
});
