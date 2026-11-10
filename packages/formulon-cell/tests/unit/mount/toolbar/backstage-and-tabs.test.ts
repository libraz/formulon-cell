import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { Spreadsheet } from '../../../../src/mount.js';
import {
  EXCEL365_STANDARD_RIBBON_TABS,
  RIBBON_TABS,
} from '../../../../src/toolbar/ribbon-model.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { stubHelpers } from './fixtures.js';

const renderedTabs = (host: HTMLElement): string[] =>
  Array.from(host.querySelectorAll<HTMLElement>('[data-ribbon-tab]')).map(
    (el) => el.dataset.ribbonTab as string,
  );

describe('Spreadsheet.mountToolbar backstage and default tabs', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    sheet.instance.host.dataset.fcPlatform = 'default';
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    sheet.dispose();
    host.remove();
  });

  it('renders the standard tab set by default and the full set when passed explicitly', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    expect(renderedTabs(host)).toEqual([...EXCEL365_STANDARD_RIBBON_TABS]);
    tb.setActiveTab('draw');
    expect(tb.getActiveTab()).toBe('home');
    tb.dispose();

    const full = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      ribbonTabs: RIBBON_TABS,
    });
    expect(renderedTabs(host)).toEqual([...RIBBON_TABS]);
    full.setActiveTab('draw');
    expect(full.getActiveTab()).toBe('draw');
    full.dispose();
  });

  it('does not open the backstage from the File tab and reports the tab change', () => {
    const onTabChange = vi.fn();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      createBackstageView: () => {
        const el = document.createElement('div');
        el.className = 'test-backstage';
        return el;
      },
      onTabChange,
    });
    host.querySelector<HTMLButtonElement>('[data-ribbon-tab="file"]')?.click();
    expect(tb.getActiveTab()).toBe('file');
    expect(tb.getBackstageOpen()).toBe(false);
    expect(host.querySelector('.test-backstage')).toBeNull();
    expect(onTabChange).toHaveBeenCalledWith('file');
    tb.dispose();
  });

  const openBackstage = (onBackstageOpenChange = vi.fn(), onTabChange = vi.fn()) => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      createBackstageView: () => {
        const el = document.createElement('div');
        el.className = 'test-backstage';
        const back = document.createElement('button');
        back.dataset.backstageAction = 'back';
        el.appendChild(back);
        return el;
      },
      onBackstageOpenChange,
      onTabChange,
    });
    tb.setActiveTab('insert');
    host.querySelector<HTMLButtonElement>('[data-ribbon-tab="file"]')?.click();
    tb.setBackstageOpen(true);
    expect(host.querySelector('.test-backstage')).not.toBeNull();
    return tb;
  };

  it('closes the backstage on Back and restores the previous tab', () => {
    const onOpen = vi.fn();
    const onTab = vi.fn();
    const tb = openBackstage(onOpen, onTab);
    host.querySelector<HTMLButtonElement>('[data-backstage-action="back"]')?.click();
    expect(tb.getBackstageOpen()).toBe(false);
    expect(host.querySelector('.test-backstage')).toBeNull();
    expect(tb.getActiveTab()).toBe('insert');
    expect(onOpen).toHaveBeenLastCalledWith(false);
    expect(onTab).toHaveBeenLastCalledWith('insert');
    tb.dispose();
  });

  it('closes the backstage on Escape', () => {
    const tb = openBackstage();
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(tb.getBackstageOpen()).toBe(false);
    expect(tb.getActiveTab()).toBe('insert');
    tb.dispose();
  });

  it('scrolls only its own tab strip to show the active tab, never the page around it', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    const strip = host.querySelector<HTMLElement>('.fc-tb__ribbon-tabs');
    const tab = host.querySelector<HTMLElement>('[data-ribbon-tab="view"]');
    if (!strip || !tab) throw new Error('missing tab strip');
    const rect = (left: number, width: number) =>
      ({ left, right: left + width, top: 0, bottom: 24, width, height: 24 }) as DOMRect;
    Object.defineProperty(strip, 'scrollWidth', { get: () => 400 });
    Object.defineProperty(strip, 'clientWidth', { get: () => 100 });
    Object.defineProperty(strip, 'offsetWidth', { get: () => 100 });
    strip.getBoundingClientRect = () => rect(0, 100);
    tab.getBoundingClientRect = () => rect(150 - strip.scrollLeft, 50);
    const pageScroll = vi.spyOn(HTMLElement.prototype, 'scrollIntoView');
    tab.click();
    expect(pageScroll).not.toHaveBeenCalled();
    expect(strip.scrollLeft).toBe(100);
    pageScroll.mockRestore();
    tb.dispose();
  });
});
