import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { mutators } from '../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

/**
 * Integration: swap locale or theme while the spreadsheet has live state.
 * Verifies the swap doesn't drop cells/selection/extras and that the
 * exposed primitives (instance.i18n + setTheme) actually flip the
 * locale/theme. Mirrors the demo apps' user flow of toggling themes mid-edit.
 */
describe('integration: locale + theme swap mid-edit', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    const { workbook, instance } = sheet;
    workbook.setText({ sheet: 0, row: 0, col: 0 }, 'alpha');
    workbook.setNumber({ sheet: 0, row: 0, col: 1 }, 42);
    workbook.recalc();
    mutators.replaceCells(instance.store, workbook.cells(0));
    mutators.setRange(instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
  });

  afterEach(() => sheet.dispose());

  it('i18n.setLocale flips the active locale without throwing or clearing state', () => {
    const { instance } = sheet;
    expect(instance.i18n.locale).toBe('en');

    instance.i18n.setLocale('ja');
    expect(instance.i18n.locale).toBe('ja');

    // Selection survives the swap.
    expect(instance.store.getState().selection.range).toEqual({
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 0,
      c1: 1,
    });
    // Cells survive.
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'alpha',
    });
  });

  it('i18n.setLocale fires localeChange subscribers (notification semantics)', () => {
    const { instance } = sheet;
    let count = 0;
    const unsub = instance.i18n.subscribe(() => {
      count += 1;
    });
    instance.i18n.setLocale('ja');
    expect(count).toBe(1);
    instance.i18n.setLocale('en');
    expect(count).toBe(2);
    unsub();
  });

  it('setTheme writes data-fc-theme on the host element', () => {
    const { instance, host } = sheet;
    instance.setTheme('ink');
    expect(host.dataset.fcTheme).toBe('ink');
    instance.setTheme('paper');
    expect(host.dataset.fcTheme).toBe('paper');
  });

  it('setTheme preserves cells and selection state', () => {
    const { instance, workbook } = sheet;
    instance.setTheme('ink');
    expect(workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'number',
      value: 42,
    });
    expect(instance.store.getState().selection.range).toEqual({
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 0,
      c1: 1,
    });
  });

  it('locale + theme swap in the same flow do not interact', () => {
    const { instance, host } = sheet;
    instance.i18n.setLocale('ja');
    instance.setTheme('ink');
    expect(instance.i18n.locale).toBe('ja');
    expect(host.dataset.fcTheme).toBe('ink');
  });

  it('locale swap during an active in-cell editor does not throw', () => {
    const { instance } = sheet;
    // Simulate "user is typing" by flipping editor mode on the store.
    mutators.setEditor(instance.store, { kind: 'edit', raw: '', caret: 0 });
    expect(() => instance.i18n.setLocale('ja')).not.toThrow();
  });
});
