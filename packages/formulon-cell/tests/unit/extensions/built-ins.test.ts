import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
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
  illustrations,
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

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

describe('built-in extension factories', () => {
  it('exposes replaceable factories for the full spreadsheet chrome', () => {
    expect(quickAnalysis().id).toBe('quickAnalysis');
    expect(viewToolbar().id).toBe('viewToolbar');
    expect(workbookObjects().id).toBe('workbookObjects');
    expect(charts().id).toBe('charts');
    expect(illustrations().id).toBe('illustrations');
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
        illustrations(),
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
      'illustrations',
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
      'illustrations',
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

  it('passes localized overlay labels into chart and illustration built-ins', () => {
    const source = readFileSync(join(root, 'src/extensions/built-ins.ts'), 'utf8');

    expect(source).toContain('labels: ctx.i18n.strings.sessionCharts');
    expect(source).toContain('pictureLabel: ctx.i18n.strings.ribbon.pictures');
    expect(source).toContain('shapeLabel: ctx.i18n.strings.ribbon.shapes');
    expect(source).toContain('resizeLabel: ctx.i18n.strings.sessionCharts.resize');
  });
});
