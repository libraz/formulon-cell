import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { setProtectedSheet } from '../../../../src/commands/protection.js';
import { attachRibbonInteractionPolicy } from '../../../../src/mount/toolbar-interaction-policy.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { stubHelpers } from './fixtures.js';

// Bold follows Format cells, as the write does; the formula auditing and pivot refresh never come back.
const REFUSED = [
  'sortAsc',
  'sortDesc',
  'precedents',
  'dependents',
  'clearArrows',
  'pivotRefresh',
  'bold',
];
const KEPT = ['copy'];

describe('ribbon projection of sheet protection without a host policy', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
  });
  afterEach(() => sheet.dispose());

  it('greys the refused commands out and re-enables them when protection lifts', () => {
    const { store } = sheet.instance;
    const host = document.createElement('div');
    for (const id of [...REFUSED, ...KEPT]) {
      const button = document.createElement('button');
      button.dataset.ribbonCommand = id;
      host.appendChild(button);
    }
    const projection = attachRibbonInteractionPolicy(host, {
      getInstance: () => sheet.instance,
      closeDynamicDropdowns: () => undefined,
    });
    const disabled = (): string[] =>
      [...host.querySelectorAll<HTMLButtonElement>('button')]
        .filter((button) => button.disabled)
        .map((button) => button.dataset.ribbonCommand ?? '');
    try {
      expect(sheet.instance.commands.policy).toBeUndefined();
      projection.project();
      expect(disabled()).toEqual([]);

      setProtectedSheet(store, 0, true);
      projection.project();
      expect(disabled()).toEqual(REFUSED);

      setProtectedSheet(store, 0, true, { permissions: { sort: true, formatCells: true } });
      projection.project();
      expect(disabled()).toEqual(['precedents', 'dependents', 'clearArrows', 'pivotRefresh']);

      setProtectedSheet(store, 0, false);
      projection.project();
      expect(disabled()).toEqual([]);

      setProtectedSheet(store, 0, true);
      projection.project();
      expect(disabled()).toEqual(REFUSED);
      sheet.workbook.addSheet('Second');
      mutators.setSheetIndex(store, 1);
      projection.project();
      expect(disabled()).toEqual([]);
    } finally {
      projection.detach();
    }
  });

  it('keeps a button another projector disabled under its own reason', () => {
    const host = document.createElement('div');
    const button = document.createElement('button');
    button.dataset.ribbonCommand = 'sortAsc';
    button.disabled = true;
    button.setAttribute('aria-description', 'needs data');
    host.appendChild(button);
    const projection = attachRibbonInteractionPolicy(host, {
      getInstance: () => sheet.instance,
      closeDynamicDropdowns: () => undefined,
    });
    try {
      setProtectedSheet(sheet.instance.store, 0, true);
      projection.project();
      expect(button.getAttribute('aria-description')).toBe('needs data');
      setProtectedSheet(sheet.instance.store, 0, false);
      projection.project();
      expect(button.disabled).toBe(true);
    } finally {
      projection.detach();
    }
  });

  it('follows protection on the mounted ribbon as the store changes', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const toolbar = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    const sort = (): HTMLButtonElement | null =>
      host.querySelector<HTMLButtonElement>('[data-ribbon-command="sortFilterHome"]');
    try {
      expect(sort()).not.toBeNull();
      expect(sort()?.disabled).toBe(false);
      setProtectedSheet(sheet.instance.store, 0, true);
      expect(sort()?.disabled).toBe(true);
      setProtectedSheet(sheet.instance.store, 0, false);
      expect(sort()?.disabled).toBe(false);
      setProtectedSheet(sheet.instance.store, 0, true);
      expect(sort()?.disabled).toBe(true);
      sheet.workbook.addSheet('Second');
      mutators.setSheetIndex(sheet.instance.store, 1);
      expect(sort()?.disabled).toBe(false);
    } finally {
      toolbar.dispose();
      host.remove();
    }
  });
});
