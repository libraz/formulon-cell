import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { mutators } from '../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

/**
 * Integration: mount the format dialog and walk the open → DOM-visible → close
 * lifecycle. The dialog is one of the heaviest pieces of chrome — it touches
 * model/state/view/dom across four files since the recent refactor. This pins
 * the smoke surface so the controller wiring doesn't silently break.
 *
 * Selector note: many dialog overlays in the codebase share `fc-fmtdlg` as the
 * base class. The format dialog is the one whose className is *only*
 * `fc-fmtdlg`; others compose `fc-fmtdlg fc-conddlg`, `fc-fmtdlg fc-hldlg`,
 * etc. We use `[class="fc-fmtdlg"]` so the test scopes to the format dialog
 * deterministically.
 */
const FORMAT_OVERLAY_SELECTOR = '[class="fc-fmtdlg"]';
type FormatDialogHandle = { open: () => void; close: () => void };
const formatDialogHandle = (sheet: MountedStubSheet): FormatDialogHandle =>
  sheet.instance.features.formatDialog as unknown as FormatDialogHandle;

describe('integration: format dialog flow', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
  });

  afterEach(() => {
    sheet.dispose();
  });

  it('exposes formatDialog through instance.features under presets.full()', () => {
    expect(sheet.instance.features.formatDialog).toBeTruthy();
  });

  it('renders the dialog overlay hidden by default', () => {
    const overlay = sheet.host.querySelector(FORMAT_OVERLAY_SELECTOR);
    expect(overlay).not.toBeNull();
    expect((overlay as HTMLElement).hidden).toBe(true);
  });

  it('open() flips the overlay visible and shows the panel', () => {
    const handle = formatDialogHandle(sheet);
    handle.open();
    const overlay = sheet.host.querySelector(FORMAT_OVERLAY_SELECTOR) as HTMLElement;
    expect(overlay.hidden).toBe(false);
    expect(overlay.querySelector('.fc-fmtdlg__panel')).not.toBeNull();
  });

  it('close() restores the hidden state', () => {
    const handle = formatDialogHandle(sheet);
    handle.open();
    handle.close();
    const overlay = sheet.host.querySelector(FORMAT_OVERLAY_SELECTOR) as HTMLElement;
    expect(overlay.hidden).toBe(true);
  });

  it('open() works while a selection is set — does not throw and renders the panel', () => {
    const { instance, workbook } = sheet;
    workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1234.5);
    mutators.replaceCells(instance.store, workbook.cells(0));
    mutators.setActive(instance.store, { sheet: 0, row: 0, col: 0 });

    const handle = formatDialogHandle(sheet);
    expect(() => handle.open()).not.toThrow();
    const overlay = sheet.host.querySelector(FORMAT_OVERLAY_SELECTOR) as HTMLElement;
    expect(overlay.hidden).toBe(false);
  });
});

describe('integration: format dialog opt-out', () => {
  it('features.formatDialog is undefined when the flag is off', async () => {
    const opt = await mountStubSheet({ features: { formatDialog: false } });
    expect(opt.instance.features.formatDialog).toBeFalsy();
    expect(opt.host.querySelector(FORMAT_OVERLAY_SELECTOR)).toBeNull();
    opt.dispose();
  });
});
