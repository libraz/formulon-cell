import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { Spreadsheet } from '../../../src/mount.js';
import type { RibbonRenderHelpers } from '../../../src/toolbar/ribbon/render-ribbon.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

// Minimal helpers stub: enough for the renderer to emit a shell, no real
// dropdown DOM. The toolbar still needs `createSelect/Color/Icon/makeSvg`
// because every command path may reach them.
const stubHelpers = (): RibbonRenderHelpers => ({
  createSelect: () => document.createElement('div'),
  createColor: () => document.createElement('div'),
  createIcon: () => null,
  makeSvg: () => document.createElementNS('http://www.w3.org/2000/svg', 'svg'),
  chevronPath: 'M0 0',
});

describe('Spreadsheet.mountToolbar', () => {
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

  it('renders the ribbon shell and returns an imperative instance', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });

    expect(tb.host).toBe(host);
    expect(tb.instance).toBe(sheet.instance);
    expect(host.querySelector('.demo__ribbon-shell')).toBeTruthy();
    expect(tb.getActiveTab()).toBe('home');
    expect(tb.getCollapsed()).toBe(false);
    expect(tb.getFormulaBarVisible()).toBe(true);
    expect(tb.getTheme()).toBe('light');

    tb.dispose();
    expect(host.children.length).toBe(0);
  });

  it('dispatches ribbon commands and fires onCommand', () => {
    const onCommand = vi.fn();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      onCommand,
    });

    // 'undoHome' is a built-in command handled by core (no hooks needed). It
    // returns true even when the history is empty — `undo()` returns false
    // but the dispatcher still claims the click.
    const applied = tb.applyCommand('undoHome');
    expect(applied).toBe(true);
    expect(onCommand).toHaveBeenCalledWith('undoHome', true);

    // Unknown ids fall through.
    const unknown = tb.applyCommand('not-a-real-command');
    expect(unknown).toBe(false);
    expect(onCommand).toHaveBeenLastCalledWith('not-a-real-command', false);

    tb.dispose();
  });

  it('clicks on ribbon tabs switch the active tab and rerender', () => {
    const onTabChange = vi.fn();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      onTabChange,
    });

    const insertTab = host.querySelector<HTMLButtonElement>('[data-ribbon-tab="insert"]');
    expect(insertTab).toBeTruthy();
    insertTab?.click();

    expect(tb.getActiveTab()).toBe('insert');
    expect(onTabChange).toHaveBeenCalledWith('insert');

    tb.dispose();
  });

  it('routes hook calls into opts.hooks when present', () => {
    const copy = vi.fn();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      hooks: { clipboard: { copy, cut: vi.fn(), paste: vi.fn() } },
    });

    tb.applyCommand('copy');
    expect(copy).toHaveBeenCalledTimes(1);

    tb.dispose();
  });

  it('dispose detaches the click listener and store subscription', () => {
    const onCommand = vi.fn();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      onCommand,
    });

    tb.dispose();

    const tabBtn = document.createElement('button');
    tabBtn.dataset.ribbonTab = 'insert';
    host.appendChild(tabBtn);
    tabBtn.click();

    // After dispose, the click listener is gone so the active tab doesn't change.
    expect(tb.getActiveTab()).toBe('home');
    expect(onCommand).not.toHaveBeenCalled();
  });
});
