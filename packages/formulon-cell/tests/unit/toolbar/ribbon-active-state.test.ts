import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { setAlign, toggleBold, toggleItalic } from '../../../src/commands/format.js';
import {
  setMarginPreset,
  setPageOrientation,
  setPaperSize,
} from '../../../src/commands/page-setup.js';
import { setFreezePanes } from '../../../src/commands/structure.js';
import { mutators } from '../../../src/store/store.js';
import {
  BORDER_PRESETS,
  BORDER_STYLES,
  EMPTY_ACTIVE_STATE,
  projectActiveState,
} from '../../../src/toolbar/ribbon-active-state.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

describe('toolbar/ribbon-active-state', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
  });

  afterEach(() => {
    sheet.dispose();
  });

  it('exposes a fully-defined empty default that mirrors a freshly mounted sheet', () => {
    const fresh = projectActiveState(sheet.instance);
    expect(EMPTY_ACTIVE_STATE.bold).toBe(false);
    expect(EMPTY_ACTIVE_STATE.italic).toBe(false);
    expect(EMPTY_ACTIVE_STATE.fontFamily).toBe('Aptos');
    expect(EMPTY_ACTIVE_STATE.fontSize).toBe(11);
    expect(EMPTY_ACTIVE_STATE.zoom).toBe(1);
    expect(EMPTY_ACTIVE_STATE.pageOrientation).toBe('portrait');
    expect(EMPTY_ACTIVE_STATE.paperSize).toBe('A4');
    expect(EMPTY_ACTIVE_STATE.marginPreset).toBe('normal');

    expect(fresh.bold).toBe(EMPTY_ACTIVE_STATE.bold);
    expect(fresh.italic).toBe(EMPTY_ACTIVE_STATE.italic);
    expect(fresh.zoom).toBe(EMPTY_ACTIVE_STATE.zoom);
    expect(fresh.pageOrientation).toBe(EMPTY_ACTIVE_STATE.pageOrientation);
    expect(fresh.paperSize).toBe(EMPTY_ACTIVE_STATE.paperSize);
    expect(fresh.marginPreset).toBe(EMPTY_ACTIVE_STATE.marginPreset);
  });

  it('reflects toggleBold / toggleItalic on the active cell', () => {
    const { instance } = sheet;
    mutators.setActive(instance.store, { sheet: 0, row: 0, col: 0 });

    toggleBold(instance.store.getState(), instance.store);
    expect(projectActiveState(instance).bold).toBe(true);

    toggleItalic(instance.store.getState(), instance.store);
    expect(projectActiveState(instance).italic).toBe(true);

    toggleBold(instance.store.getState(), instance.store);
    expect(projectActiveState(instance).bold).toBe(false);
    expect(projectActiveState(instance).italic).toBe(true);
  });

  it('mirrors alignment as exclusive booleans', () => {
    const { instance } = sheet;
    mutators.setActive(instance.store, { sheet: 0, row: 0, col: 0 });

    setAlign(instance.store.getState(), instance.store, 'center');
    let s = projectActiveState(instance);
    expect(s.alignCenter).toBe(true);
    expect(s.alignLeft).toBe(false);
    expect(s.alignRight).toBe(false);

    setAlign(instance.store.getState(), instance.store, 'right');
    s = projectActiveState(instance);
    expect(s.alignCenter).toBe(false);
    expect(s.alignRight).toBe(true);
  });

  it('reports frozen=true when freezeRows or freezeCols are set', () => {
    const { instance } = sheet;
    expect(projectActiveState(instance).frozen).toBe(false);

    setFreezePanes(instance.store, null, 1, 0);
    expect(projectActiveState(instance).frozen).toBe(true);

    setFreezePanes(instance.store, null, 0, 0);
    expect(projectActiveState(instance).frozen).toBe(false);

    setFreezePanes(instance.store, null, 0, 2);
    expect(projectActiveState(instance).frozen).toBe(true);
  });

  it('mirrors page setup (orientation, paper size, margin preset)', () => {
    const { instance } = sheet;
    setPageOrientation(instance.store, 0, 'landscape');
    setPaperSize(instance.store, 0, 'letter');
    setMarginPreset(instance.store, 0, 'wide');

    const s = projectActiveState(instance);
    expect(s.pageOrientation).toBe('landscape');
    expect(s.paperSize).toBe('letter');
    expect(s.marginPreset).toBe('wide');
  });

  it('exposes border presets/styles as label/value tuples covering all common variants', () => {
    expect(BORDER_PRESETS.map((p) => p.value)).toEqual([
      'none',
      'outline',
      'all',
      'top',
      'bottom',
      'left',
      'right',
      'doubleBottom',
    ]);
    expect(BORDER_STYLES.map((s) => s.value)).toEqual([
      'thin',
      'medium',
      'thick',
      'dashed',
      'dotted',
      'double',
    ]);
    for (const p of BORDER_PRESETS) expect(typeof p.label).toBe('string');
    for (const s of BORDER_STYLES) expect(typeof s.label).toBe('string');
  });

  it('keeps zoom in sync with the viewport zoom level', () => {
    const { instance } = sheet;
    mutators.setZoom(instance.store, 1.25);
    expect(projectActiveState(instance).zoom).toBeCloseTo(1.25);
  });
});
