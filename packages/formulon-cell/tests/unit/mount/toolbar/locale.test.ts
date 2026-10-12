import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import { toolbarMenuText } from '../../../../src/toolbar/menu-text.js';
import { toolbarText } from '../../../../src/toolbar/ribbon-model.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { stubHelpers } from './fixtures.js';

describe('Spreadsheet.mountToolbar locale updates', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    sheet.dispose();
    host.remove();
  });

  it('refreshes default ribbon and dropdown labels while preserving toolbar state', () => {
    const toolbar = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    const paste = () => host.querySelector<HTMLButtonElement>('[data-ribbon-command="paste"]');
    expect(paste()?.textContent).toContain('Paste');
    paste()?.click();
    expect(host.querySelector('#menu-paste button')?.textContent).toContain('Paste');

    toolbar.setActiveTab('insert');
    toolbar.setCollapsed(true);
    expect(toolbar.getActiveTab()).toBe('insert');
    expect(toolbar.getDisplayMode()).toBe('tabsOnly');

    sheet.instance.setLocale('ja');

    expect(toolbar.getActiveTab()).toBe('insert');
    expect(toolbar.getDisplayMode()).toBe('tabsOnly');
    expect(host.querySelector('[data-ribbon-tab="insert"]')).not.toBeNull();
    expect(host.querySelector('#menu-paste button')?.textContent).toContain('貼り付け');

    toolbar.dispose();
  });

  it('keeps explicit ribbon text overrides across locale changes', () => {
    const explicitText = { ...toolbarText('en'), disabled: 'Keep this label' };
    const toolbar = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      text: explicitText,
    });

    expect(
      host.querySelector<HTMLButtonElement>('[data-ribbon-command="helpSearch"]')?.dataset
        .ribbonDisabledReason,
    ).toBe('Keep this label');

    sheet.instance.setLocale('ja');

    expect(
      host.querySelector<HTMLButtonElement>('[data-ribbon-command="helpSearch"]')?.dataset
        .ribbonDisabledReason,
    ).toBe('Keep this label');
    toolbar.dispose();
  });

  it('unsubscribes from locale changes when disposed', () => {
    const toolbar = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    toolbar.dispose();

    expect(() => sheet.instance.setLocale('ja')).not.toThrow();
    expect(host.children).toHaveLength(0);
  });

  it('sets the toolbar language independently and restores the embedding attribute', () => {
    host.lang = 'it';
    const toolbar = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    expect(host.lang).toBe('en');
    sheet.instance.setLocale('de');
    expect(host.lang).toBe('de');
    toolbar.dispose();
    expect(host.lang).toBe('it');
  });

  it('keeps an explicitly selected toolbar language across sheet locale changes', () => {
    const toolbar = Spreadsheet.mountToolbar(host, sheet.instance, {
      lang: 'ja',
      helpers: stubHelpers(),
    });
    expect(host.lang).toBe('ja');
    sheet.instance.setLocale('fr');
    expect(host.lang).toBe('ja');
    toolbar.dispose();
    expect(host.hasAttribute('lang')).toBe(false);
  });

  it('uses explicit toolbar background text while the sheet locale differs', async () => {
    sheet.instance.setLocale('ja');
    const toolbar = Spreadsheet.mountToolbar(host, sheet.instance, {
      lang: 'en',
      text: toolbarText('en'),
      menuText: toolbarMenuText('en'),
      helpers: stubHelpers(),
    });
    toolbar.setActiveTab('pageLayout');

    const backgroundButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="sheetBackground"]',
    );
    expect(backgroundButton?.textContent).toContain('Background');
    expect(backgroundButton?.getAttribute('aria-label')).toBe('Background');

    mutators.setSheetBackgroundImage(
      sheet.instance.store,
      0,
      'https://example.test/explicit-toolbar-background.png',
    );
    await Promise.resolve();
    expect(backgroundButton?.textContent).toContain('Delete Background');
    expect(backgroundButton?.getAttribute('aria-label')).toBe('Delete Background');

    mutators.setSheetBackgroundImage(sheet.instance.store, 0, undefined);
    await Promise.resolve();
    expect(backgroundButton?.textContent).toContain('Background');
    expect(backgroundButton?.getAttribute('aria-label')).toBe('Background');

    toolbar.dispose();
  });

  it('refreshes dynamic background labels when the toolbar follows the sheet locale', async () => {
    const toolbar = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
    });
    toolbar.setActiveTab('pageLayout');
    const backgroundButton = () =>
      host.querySelector<HTMLButtonElement>('[data-ribbon-command="sheetBackground"]');
    expect(backgroundButton()?.textContent).toContain(toolbarText('en').background);

    sheet.instance.setLocale('ja');
    await Promise.resolve();
    expect(backgroundButton()?.textContent).toContain(toolbarText('ja').background);

    mutators.setSheetBackgroundImage(sheet.instance.store, 0, 'https://example.test/ja-bg.png');
    await Promise.resolve();
    expect(backgroundButton()?.textContent).toContain(toolbarMenuText('ja').sheetBackgroundClear);

    toolbar.dispose();
  });
});
