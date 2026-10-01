import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { addrKey } from '../../../src/engine/address.js';
import { Spreadsheet } from '../../../src/mount.js';
import type { RibbonRenderHelpers } from '../../../src/toolbar/ribbon/render-ribbon.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

const stubHelpers = (): RibbonRenderHelpers => ({
  createSelect: () => document.createElement('div'),
  createColor: () => document.createElement('div'),
  createIcon: () => null,
  makeSvg: (_viewBox, _pathData, className) => {
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.setAttribute('class', className);
    return svg;
  },
  chevronPath: 'M0 0',
});

const seedMergeRange = (sheet: MountedStubSheet): void => {
  sheet.instance.store.setState((state) => ({
    ...state,
    selection: {
      ...state.selection,
      active: { sheet: 0, row: 0, col: 0 },
      range: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
    },
  }));
};

const clickChevron = (button: HTMLButtonElement, nestedPath = false): void => {
  const chevron = button.querySelector<SVGSVGElement>('.fc-tb__rb-split-chevron');
  expect(chevron).toBeTruthy();
  if (!chevron) return;
  const target = nestedPath
    ? (() => {
        const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
        chevron.appendChild(path);
        return path;
      })()
    : chevron;
  target.dispatchEvent(new MouseEvent('click', { bubbles: true }));
};

describe('Spreadsheet.mountToolbar split-button activation', () => {
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

  it('opens the merge menu from a bubbling chevron click without merging', () => {
    seedMergeRange(sheet);
    const toolbar = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const button = host.querySelector<HTMLButtonElement>('[data-ribbon-command="merge"]');
    const menu = host.querySelector<HTMLElement>('#menu-merge');
    expect(button).toBeTruthy();
    expect(menu?.hidden).toBe(true);

    clickChevron(button as HTMLButtonElement);

    expect(menu?.hidden).toBe(false);
    expect(button?.getAttribute('aria-expanded')).toBe('true');
    expect(sheet.instance.store.getState().merges.byAnchor.size).toBe(0);
    toolbar.dispose();
  });

  it('opens the merge menu when the bubbling target is nested inside the chevron', () => {
    seedMergeRange(sheet);
    const toolbar = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const button = host.querySelector<HTMLButtonElement>('[data-ribbon-command="merge"]');
    const menu = host.querySelector<HTMLElement>('#menu-merge');
    expect(button).toBeTruthy();

    clickChevron(button as HTMLButtonElement, true);

    expect(menu?.hidden).toBe(false);
    expect(sheet.instance.store.getState().merges.byAnchor.size).toBe(0);
    toolbar.dispose();
  });

  it('keeps the primary face action on merge split buttons', () => {
    seedMergeRange(sheet);
    const toolbar = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const button = host.querySelector<HTMLButtonElement>('[data-ribbon-command="merge"]');
    expect(button).toBeTruthy();

    button?.click();

    expect(
      sheet.instance.store.getState().merges.byAnchor.get(addrKey({ sheet: 0, row: 0, col: 0 })),
    ).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    expect(host.querySelector<HTMLElement>('#menu-merge')?.hidden).toBe(true);
    toolbar.dispose();
  });

  it('opens Paste options from the chevron without invoking paste', () => {
    const paste = vi.fn();
    const toolbar = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
      hooks: { clipboard: { paste } },
    });
    const button = host.querySelector<HTMLButtonElement>('[data-ribbon-command="paste"]');
    const menu = host.querySelector<HTMLElement>('#menu-paste');
    expect(button).toBeTruthy();

    clickChevron(button as HTMLButtonElement);

    expect(menu?.hidden).toBe(false);
    expect(paste).not.toHaveBeenCalled();
    toolbar.dispose();
  });

  it('opens a static split menu when dynamic dropdowns are not enabled', () => {
    seedMergeRange(sheet);
    const toolbar = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
    });
    const button = host.querySelector<HTMLButtonElement>('[data-ribbon-command="merge"]');
    const menu = host.querySelector<HTMLElement>('#menu-merge');
    expect(button).toBeTruthy();
    expect(menu).toBeTruthy();

    clickChevron(button as HTMLButtonElement);

    expect(menu?.hidden).toBe(false);
    expect(button?.getAttribute('aria-expanded')).toBe('true');
    expect(sheet.instance.store.getState().merges.byAnchor.size).toBe(0);
    toolbar.dispose();
  });

  it('does not open or dispatch a disabled split button', () => {
    seedMergeRange(sheet);
    const toolbar = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const button = host.querySelector<HTMLButtonElement>('[data-ribbon-command="merge"]');
    const menu = host.querySelector<HTMLElement>('#menu-merge');
    expect(button).toBeTruthy();
    if (!button) return;
    button.disabled = true;

    clickChevron(button);

    expect(menu?.hidden).toBe(true);
    expect(button.getAttribute('aria-expanded')).toBe('false');
    expect(sheet.instance.store.getState().merges.byAnchor.size).toBe(0);
    toolbar.dispose();
  });
});
