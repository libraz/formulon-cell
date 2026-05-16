import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { formatAsTable } from '../../../src/commands/format-as-table.js';
import { applyCellStyle } from '../../../src/commands/cell-styles.js';
import {
  setAlign,
  setNumFmt,
  setRotation,
  setVAlign,
  toggleBold,
  toggleItalic,
  toggleWrap,
} from '../../../src/commands/format.js';
import {
  setFitToPages,
  setMarginPreset,
  setPageOrientation,
  setPageScale,
  setPaperSize,
  setPrintGridlines,
  setPrintHeadings,
} from '../../../src/commands/page-setup.js';
import { setFreezePanes } from '../../../src/commands/structure.js';
import { dictionaries } from '../../../src/i18n/strings.js';
import { mutators } from '../../../src/store/store.js';
import {
  BORDER_PRESETS,
  BORDER_STYLES,
  EMPTY_ACTIVE_STATE,
  localizeBorderPresets,
  localizeBorderStyles,
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
    expect(EMPTY_ACTIVE_STATE.gridlinesVisible).toBe(true);
    expect(EMPTY_ACTIVE_STATE.headingsVisible).toBe(true);
    expect(EMPTY_ACTIVE_STATE.formulasVisible).toBe(false);
    expect(EMPTY_ACTIVE_STATE.r1c1).toBe(false);
    expect(EMPTY_ACTIVE_STATE.vAlignBottom).toBe(true);
    expect(EMPTY_ACTIVE_STATE.wrapText).toBe(false);
    expect(EMPTY_ACTIVE_STATE.merged).toBe(false);
    expect(EMPTY_ACTIVE_STATE.mergeCenter).toBe(false);
    expect(EMPTY_ACTIVE_STATE.conditionalFormatting).toBe(false);
    expect(EMPTY_ACTIVE_STATE.formatAsTable).toBe(false);
    expect(EMPTY_ACTIVE_STATE.cellStyle).toBeNull();
    expect(EMPTY_ACTIVE_STATE.textOrientation).toBe('horizontalText');
    expect(EMPTY_ACTIVE_STATE.commaStyle).toBe(false);
    expect(EMPTY_ACTIVE_STATE.marginPreset).toBe('normal');

    expect(fresh.bold).toBe(EMPTY_ACTIVE_STATE.bold);
    expect(fresh.italic).toBe(EMPTY_ACTIVE_STATE.italic);
    expect(fresh.zoom).toBe(EMPTY_ACTIVE_STATE.zoom);
    expect(fresh.pageOrientation).toBe(EMPTY_ACTIVE_STATE.pageOrientation);
    expect(fresh.paperSize).toBe(EMPTY_ACTIVE_STATE.paperSize);
    expect(fresh.gridlinesVisible).toBe(EMPTY_ACTIVE_STATE.gridlinesVisible);
    expect(fresh.headingsVisible).toBe(EMPTY_ACTIVE_STATE.headingsVisible);
    expect(fresh.formulasVisible).toBe(EMPTY_ACTIVE_STATE.formulasVisible);
    expect(fresh.r1c1).toBe(EMPTY_ACTIVE_STATE.r1c1);
    expect(fresh.vAlignBottom).toBe(EMPTY_ACTIVE_STATE.vAlignBottom);
    expect(fresh.wrapText).toBe(EMPTY_ACTIVE_STATE.wrapText);
    expect(fresh.merged).toBe(EMPTY_ACTIVE_STATE.merged);
    expect(fresh.mergeCenter).toBe(EMPTY_ACTIVE_STATE.mergeCenter);
    expect(fresh.conditionalFormatting).toBe(EMPTY_ACTIVE_STATE.conditionalFormatting);
    expect(fresh.formatAsTable).toBe(EMPTY_ACTIVE_STATE.formatAsTable);
    expect(fresh.cellStyle).toBe(EMPTY_ACTIVE_STATE.cellStyle);
    expect(fresh.textOrientation).toBe(EMPTY_ACTIVE_STATE.textOrientation);
    expect(fresh.commaStyle).toBe(EMPTY_ACTIVE_STATE.commaStyle);
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

  it('mirrors comma style for fixed formats with thousands separators', () => {
    const { instance } = sheet;
    mutators.setActive(instance.store, { sheet: 0, row: 0, col: 0 });

    setNumFmt(instance.store.getState(), instance.store, {
      kind: 'fixed',
      decimals: 2,
      thousands: true,
    });
    expect(projectActiveState(instance).commaStyle).toBe(true);

    setNumFmt(instance.store.getState(), instance.store, {
      kind: 'fixed',
      decimals: 2,
      thousands: false,
    });
    expect(projectActiveState(instance).commaStyle).toBe(false);
  });

  it('mirrors vertical alignment and wrap text state', () => {
    const { instance } = sheet;
    mutators.setActive(instance.store, { sheet: 0, row: 0, col: 0 });

    let state = projectActiveState(instance);
    expect(state.vAlignBottom).toBe(true);
    expect(state.vAlignTop).toBe(false);
    expect(state.vAlignMiddle).toBe(false);
    expect(state.wrapText).toBe(false);

    setVAlign(instance.store.getState(), instance.store, 'top');
    toggleWrap(instance.store.getState(), instance.store);

    state = projectActiveState(instance);
    expect(state.vAlignTop).toBe(true);
    expect(state.vAlignMiddle).toBe(false);
    expect(state.vAlignBottom).toBe(false);
    expect(state.wrapText).toBe(true);

    setVAlign(instance.store.getState(), instance.store, 'middle');
    state = projectActiveState(instance);
    expect(state.vAlignTop).toBe(false);
    expect(state.vAlignMiddle).toBe(true);
    expect(state.vAlignBottom).toBe(false);
  });

  it('mirrors text orientation from cell rotation', () => {
    const { instance } = sheet;
    mutators.setActive(instance.store, { sheet: 0, row: 0, col: 0 });

    expect(projectActiveState(instance).textOrientation).toBe('horizontalText');

    setRotation(instance.store.getState(), instance.store, 45);
    expect(projectActiveState(instance).textOrientation).toBe('angleCounterclockwise');

    setRotation(instance.store.getState(), instance.store, -45);
    expect(projectActiveState(instance).textOrientation).toBe('angleClockwise');

    setRotation(instance.store.getState(), instance.store, 90);
    expect(projectActiveState(instance).textOrientation).toBe('rotateTextUp');

    setRotation(instance.store.getState(), instance.store, -90);
    expect(projectActiveState(instance).textOrientation).toBe('rotateTextDown');

    setRotation(instance.store.getState(), instance.store, 0);
    expect(projectActiveState(instance).textOrientation).toBe('horizontalText');
  });

  it('mirrors merged and merge-center state from the active cell', () => {
    const { instance } = sheet;
    mutators.setActive(instance.store, { sheet: 0, row: 1, col: 1 });

    mutators.mergeRange(instance.store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 2 });
    let state = projectActiveState(instance);
    expect(state.merged).toBe(true);
    expect(state.mergeCenter).toBe(false);

    mutators.setCellFormat(instance.store, { sheet: 0, row: 1, col: 1 }, { align: 'center' });
    state = projectActiveState(instance);
    expect(state.merged).toBe(true);
    expect(state.mergeCenter).toBe(true);

    mutators.setActive(instance.store, { sheet: 0, row: 1, col: 2 });
    state = projectActiveState(instance);
    expect(state.merged).toBe(true);
    expect(state.mergeCenter).toBe(true);
    expect(state.alignCenter).toBe(true);
  });

  it('marks conditional formatting active when the selected range intersects a rule', () => {
    const { instance } = sheet;
    mutators.setRange(instance.store, { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 });

    expect(projectActiveState(instance).conditionalFormatting).toBe(false);

    mutators.addConditionalRule(instance.store, {
      kind: 'data-bar',
      range: { sheet: 0, r0: 2, c0: 2, r1: 3, c1: 3 },
      color: '#70ad47',
      showValue: true,
    });
    expect(projectActiveState(instance).conditionalFormatting).toBe(true);

    mutators.setRange(instance.store, { sheet: 0, r0: 4, c0: 4, r1: 4, c1: 4 });
    expect(projectActiveState(instance).conditionalFormatting).toBe(false);
  });

  it('marks format-as-table active when the active cell is inside a table overlay', () => {
    const { instance } = sheet;
    formatAsTable(instance.store, { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 3 }, { style: 'medium' });

    mutators.setActive(instance.store, { sheet: 0, row: 2, col: 2 });
    expect(projectActiveState(instance).formatAsTable).toBe(true);

    mutators.setActive(instance.store, { sheet: 0, row: 4, col: 4 });
    expect(projectActiveState(instance).formatAsTable).toBe(false);
  });

  it('mirrors the active cell named style id', () => {
    const { instance } = sheet;
    applyCellStyle(instance.store, null, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 }, 'good');

    mutators.setActive(instance.store, { sheet: 0, row: 1, col: 1 });
    expect(projectActiveState(instance).cellStyle).toBe('good');

    applyCellStyle(instance.store, null, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 }, 'normal');
    expect(projectActiveState(instance).cellStyle).toBeNull();
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

  it('mirrors page setup (orientation, paper size, margin preset, scale, print options)', () => {
    const { instance } = sheet;
    setPageOrientation(instance.store, 0, 'landscape');
    setPaperSize(instance.store, 0, 'letter');
    setMarginPreset(instance.store, 0, 'wide');
    setPageScale(instance.store, 0, 0.75);
    setFitToPages(instance.store, 0, 'width', 1);
    setFitToPages(instance.store, 0, 'height', 2);
    setPrintGridlines(instance.store, 0, true);
    setPrintHeadings(instance.store, 0, true);

    const s = projectActiveState(instance);
    expect(s.pageOrientation).toBe('landscape');
    expect(s.paperSize).toBe('letter');
    expect(s.marginPreset).toBe('wide');
    expect(s.pageScale).toBe(0.75);
    expect(s.fitWidth).toBe(1);
    expect(s.fitHeight).toBe(2);
    expect(s.printGridlines).toBe(true);
    expect(s.printHeadings).toBe(true);
  });

  it('exposes border presets/styles as label/value tuples covering all common variants', () => {
    expect(BORDER_PRESETS.map((p) => p.value)).toEqual([
      'none',
      'outline',
      'thickOutline',
      'all',
      'inside',
      'insideHorizontal',
      'insideVertical',
      'top',
      'bottom',
      'left',
      'right',
      'doubleBottom',
      'thickBottom',
      'topAndBottom',
      'topAndThickBottom',
      'topAndDoubleBottom',
      'diagonalDown',
      'diagonalUp',
    ]);
    expect(BORDER_STYLES.map((s) => s.value)).toEqual([
      'thin',
      'medium',
      'thick',
      'dashed',
      'dotted',
      'double',
      'hair',
      'mediumDashed',
      'dashDot',
      'mediumDashDot',
      'dashDotDot',
      'mediumDashDotDot',
      'slantDashDot',
    ]);
    for (const p of BORDER_PRESETS) expect(typeof p.label).toBe('string');
    for (const s of BORDER_STYLES) expect(typeof s.label).toBe('string');
    for (const p of BORDER_PRESETS) expect(typeof p.labelKey).toBe('string');
    for (const s of BORDER_STYLES) expect(typeof s.labelKey).toBe('string');
  });

  it('localizes border presets/styles through the shared i18n dictionary', () => {
    expect(localizeBorderPresets(dictionaries.ja.ribbon).map((p) => p.label)).toEqual(
      expect.arrayContaining(['罫線なし', '外枠', '右下がり斜め罫線']),
    );
    expect(localizeBorderStyles(dictionaries.ja.ribbon).map((s) => s.label)).toEqual(
      expect.arrayContaining(['細線', '中太破線', '斜め一点鎖線']),
    );
  });

  it('keeps zoom in sync with the viewport zoom level', () => {
    const { instance } = sheet;
    mutators.setZoom(instance.store, 1.25);
    expect(projectActiveState(instance).zoom).toBeCloseTo(1.25);
  });

  it('mirrors view toggles for gridlines, headings, formulas, and R1C1', () => {
    const { instance } = sheet;
    let s = projectActiveState(instance);
    expect(s.gridlinesVisible).toBe(true);
    expect(s.headingsVisible).toBe(true);
    expect(s.formulasVisible).toBe(false);
    expect(s.workbookView).toBe('normal');
    expect(s.r1c1).toBe(false);

    mutators.setShowGridLines(instance.store, false);
    mutators.setShowHeaders(instance.store, false);
    mutators.setShowFormulas(instance.store, true);
    mutators.setWorkbookView(instance.store, 'pageBreakPreview');
    mutators.setR1C1(instance.store, true);

    s = projectActiveState(instance);
    expect(s.gridlinesVisible).toBe(false);
    expect(s.headingsVisible).toBe(false);
    expect(s.formulasVisible).toBe(true);
    expect(s.workbookView).toBe('pageBreakPreview');
    expect(s.r1c1).toBe(true);
  });
});
