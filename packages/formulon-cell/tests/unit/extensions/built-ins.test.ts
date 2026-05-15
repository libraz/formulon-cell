import { describe, expect, it } from 'vitest';
import {
  allBuiltIns,
  charts,
  clipboard,
  commentDialog,
  conditionalDialog,
  contextMenu,
  findReplace,
  formatDialog,
  formatPainter,
  goToSpecialDialog,
  hoverComment,
  hyperlinkDialog,
  iterativeDialog,
  namedRangeDialog,
  pageSetupDialog,
  pasteSpecial,
  pivotTableDialog,
  quickAnalysis,
  slicer,
  statusBar,
  validationList,
  viewToolbar,
  watchWindow,
  wheel,
  workbookObjects,
} from '../../../src/extensions/built-ins.js';
import { ALL_FEATURE_IDS } from '../../../src/extensions/features.js';

describe('built-in extension factories', () => {
  it('exposes replaceable factories for the full spreadsheet chrome', () => {
    expect(quickAnalysis().id).toBe('quickAnalysis');
    expect(viewToolbar().id).toBe('viewToolbar');
    expect(workbookObjects().id).toBe('workbookObjects');
    expect(charts().id).toBe('charts');
    expect(pivotTableDialog().id).toBe('pivotTableDialog');
    expect(watchWindow().id).toBe('watchWindow');
    expect(slicer().id).toBe('slicer');
  });

  it('keeps factory ids aligned with feature ids', () => {
    expect(
      [
        formatPainter(),
        quickAnalysis(),
        charts(),
        pivotTableDialog(),
        statusBar(),
        workbookObjects(),
        viewToolbar(),
        hoverComment(),
        goToSpecialDialog(),
        conditionalDialog(),
        iterativeDialog(),
        pageSetupDialog(),
        namedRangeDialog(),
        hyperlinkDialog(),
        commentDialog(),
        formatDialog(),
        findReplace(),
        validationList(),
        clipboard(),
        pasteSpecial(),
        contextMenu(),
        wheel(),
      ].map((ext) => ext.id),
    ).toEqual([
      'formatPainter',
      'quickAnalysis',
      'charts',
      'pivotTableDialog',
      'statusBar',
      'workbookObjects',
      'viewToolbar',
      'hoverComment',
      'gotoSpecial',
      'conditional',
      'iterative',
      'pageSetup',
      'namedRanges',
      'hyperlink',
      'commentDialog',
      'formatDialog',
      'findReplace',
      'validation',
      'clipboard',
      'pasteSpecial',
      'contextMenu',
      'wheel',
    ]);
  });

  it('bundles default-on public built-in factories in allBuiltIns', () => {
    expect(allBuiltIns().map((ext) => ext.id)).toEqual([
      'formatPainter',
      'borderDraw',
      'quickAnalysis',
      'charts',
      'pivotTableDialog',
      'statusBar',
      'workbookObjects',
      'viewToolbar',
      'hoverComment',
      'gotoSpecial',
      'conditional',
      'iterative',
      'pageSetup',
      'namedRanges',
      'hyperlink',
      'commentDialog',
      'formatDialog',
      'findReplace',
      'validation',
      'clipboard',
      'pasteSpecial',
      'contextMenu',
      'wheel',
    ]);
  });

  it('keeps every public factory id inside FeatureId', () => {
    const featureIds = new Set<string>(ALL_FEATURE_IDS);
    const factories = [...allBuiltIns(), watchWindow(), slicer()];

    expect(factories.map((ext) => ext.id).filter((id) => !featureIds.has(id))).toEqual([]);
  });
});
