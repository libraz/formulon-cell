import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { mutators } from '../../../src/store/store.js';
import { pageScaleMenuText } from '../../../src/toolbar/menu-text.js';
import { createControlDispatch } from '../../../src/toolbar/ribbon/control-dispatch.js';
import { toolbarText } from '../../../src/toolbar/ribbon-model.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

describe('toolbar/ribbon control dispatch active format', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
  });

  afterEach(() => {
    sheet.dispose();
  });

  it('reads effective font and fill colors from the active merged format', () => {
    const { instance } = sheet;
    const dispatch = createControlDispatch({
      getInst: () => instance,
      ribbonLang: 'en',
      ribbonText: toolbarText('en'),
      pageScaleText: pageScaleMenuText('en'),
      sheetEl: instance.host,
      focusSheet: vi.fn(),
      refreshWorkbookCells: vi.fn(),
      projectFormatToolbar: vi.fn(),
    });
    const anchor = { sheet: 0, row: 1, col: 1 };
    const member = { sheet: 0, row: 1, col: 2 };

    mutators.setCellFormat(instance.store, anchor, {
      color: '#123456',
      fill: '#abcdef',
    });
    mutators.mergeRange(instance.store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 2 });
    mutators.setActive(instance.store, member);

    expect(dispatch.currentRibbonControlValue('fontColor')).toBe('#123456');
    expect(dispatch.currentRibbonControlValue('fillColor')).toBe('#abcdef');

    mutators.setPendingFormat(instance.store, {
      addr: member,
      format: { color: '#fedcba', fill: '#654321' },
    });
    expect(dispatch.currentRibbonControlValue('fontColor')).toBe('#fedcba');
    expect(dispatch.currentRibbonControlValue('fillColor')).toBe('#654321');
  });
});
