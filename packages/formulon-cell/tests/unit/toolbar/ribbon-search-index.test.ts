import { describe, expect, it } from 'vitest';
import {
  buildRibbonSearchIndex,
  queryRibbonSearchIndex,
} from '../../../src/toolbar/ribbon/search-index.js';
import {
  EXCEL365_STANDARD_RIBBON_TABS,
  OPTIONAL_RIBBON_TABS,
  ribbonActivatableSurfaceCommandIds,
} from '../../../src/toolbar/ribbon-model.js';
import { ribbonActivationForCommand } from '../../../src/toolbar/ribbon/activation.js';

describe('toolbar/ribbon/search-index', () => {
  it('builds Search/Tell me items from the shared ribbon model', () => {
    const items = buildRibbonSearchIndex('en', { tabs: EXCEL365_STANDARD_RIBBON_TABS });

    expect(items.find((item) => item.id === 'tab:file')).toMatchObject({
      kind: 'tab',
      label: 'File',
      tab: 'file',
    });
    expect(items.find((item) => item.commandId === 'pivotTableInsert')).toMatchObject({
      kind: 'command',
      label: 'PivotTable',
      tab: 'insert',
    });
    expect(items.find((item) => item.commandId === 'printArea')).toMatchObject({
      kind: 'command',
      label: 'Print Area',
      tab: 'pageLayout',
    });
    expect(items.find((item) => item.commandId === 'pivotFieldListView')).toMatchObject({
      kind: 'command',
      label: 'PivotTable Fields',
      tab: 'view',
    });
    expect(items.find((item) => item.id === 'help:helpAndTraining')).toMatchObject({
      kind: 'help',
      label: 'Help and training',
      tab: 'help',
    });
    expect(items.some((item) => item.tab === 'automate')).toBe(false);
    expect(items.some((item) => item.tab === 'acrobat')).toBe(false);
    expect(items.some((item) => item.commandId === 'script')).toBe(false);
    expect(items.some((item) => item.commandId === 'pdf')).toBe(false);
  });

  it('includes optional add-in tabs only when the caller includes that tab surface', () => {
    const items = buildRibbonSearchIndex('en', { tabs: OPTIONAL_RIBBON_TABS });

    expect(queryRibbonSearchIndex(items, 'script', 8)[0]).toMatchObject({
      commandId: 'script',
      tab: 'automate',
    });
    expect(queryRibbonSearchIndex(items, 'pdf', 8)[0]).toMatchObject({
      commandId: 'pdf',
      tab: 'acrobat',
    });
  });

  it('queries labels, hints, groups, tab names, ids, and option labels', () => {
    const items = buildRibbonSearchIndex('en', { tabs: EXCEL365_STANDARD_RIBBON_TABS });

    expect(queryRibbonSearchIndex(items, 'pivot')[0]?.commandId).toBe('pivotTableInsert');
    expect(queryRibbonSearchIndex(items, 'pivot fields')[0]?.commandId).toBe('pivotFieldListView');
    expect(queryRibbonSearchIndex(items, 'page layout')[0]?.tab).toBe('pageLayout');
    expect(queryRibbonSearchIndex(items, 'landscape')[0]?.commandId).toBe('orientationPreset');
    expect(queryRibbonSearchIndex(items, 'recalc')[0]?.commandId).toBe('recalcNow');
    expect(queryRibbonSearchIndex(items, 'support')[0]?.kind).toBe('help');
  });

  it('ranks exact and prefix matches ahead of broad keyword matches', () => {
    const items = buildRibbonSearchIndex('en', { tabs: EXCEL365_STANDARD_RIBBON_TABS });

    expect(queryRibbonSearchIndex(items, 'Insert')[0]).toMatchObject({
      kind: 'tab',
      tab: 'insert',
    });
    expect(queryRibbonSearchIndex(items, 'Print Area')[0]).toMatchObject({
      kind: 'command',
      commandId: 'printArea',
      tab: 'pageLayout',
    });
    expect(queryRibbonSearchIndex(items, 'Format as')[0]).toMatchObject({
      commandId: 'formatTableHome',
    });
    expect(queryRibbonSearchIndex(items, 'insert table')[0]).toMatchObject({
      commandId: 'formatTableInsert',
      tab: 'insert',
    });
  });

  it('keeps disabled matches below enabled exact matches when included', () => {
    const items = buildRibbonSearchIndex('en', { includeDisabled: true });
    const matches = queryRibbonSearchIndex(items, 'help', 4);

    expect(matches[0]).toMatchObject({ kind: 'tab', tab: 'help' });
    expect(matches[0]?.disabled).toBeUndefined();
    expect(matches.find((item) => item.commandId === 'helpSearch')).toMatchObject({
      disabled: true,
    });
  });

  it('matches common spreadsheet synonyms without framework-specific search code', () => {
    const items = buildRibbonSearchIndex('en', { tabs: EXCEL365_STANDARD_RIBBON_TABS });

    expect(queryRibbonSearchIndex(items, 'lock panes')[0]).toMatchObject({
      commandId: 'freeze',
      tab: 'view',
    });
    expect(queryRibbonSearchIndex(items, 'repeat rows')[0]).toMatchObject({
      commandId: 'printTitles',
      tab: 'pageLayout',
    });
    expect(queryRibbonSearchIndex(items, 'split columns')[0]).toMatchObject({
      commandId: 'textToColumns',
      tab: 'data',
    });
    expect(queryRibbonSearchIndex(items, 'dedupe')[0]).toMatchObject({
      commandId: 'removeDupes',
      tab: 'data',
    });
    expect(queryRibbonSearchIndex(items, 'field list')[0]).toMatchObject({
      commandId: 'pivotFieldListView',
      tab: 'view',
    });
    expect(queryRibbonSearchIndex(items, 'selection pane')[0]).toMatchObject({
      commandId: 'selectionPanePageLayout',
      tab: 'pageLayout',
    });
    expect(queryRibbonSearchIndex(items, 'bring to front')[0]).toMatchObject({
      commandId: 'arrangeObjectsPageLayout',
      tab: 'pageLayout',
    });
    expect(queryRibbonSearchIndex(items, 'check accessibility')[0]).toMatchObject({
      commandId: 'accessibility',
      tab: 'review',
    });
    expect(queryRibbonSearchIndex(items, 'combine cells')[0]).toMatchObject({
      commandId: 'merge',
      tab: 'home',
    });
    expect(queryRibbonSearchIndex(items, 'unmerge cells')[0]).toMatchObject({
      commandId: 'merge',
      tab: 'home',
    });
    expect(queryRibbonSearchIndex(items, 'search and select')[0]).toMatchObject({
      commandId: 'findHome',
      tab: 'home',
    });
    expect(queryRibbonSearchIndex(items, 'show formula bar')[0]).toMatchObject({
      commandId: 'viewFormulaBar',
      tab: 'view',
    });
    expect(queryRibbonSearchIndex(items, 'show gridlines')[0]).toMatchObject({
      commandId: 'pageLayoutGridlinesView',
      tab: 'pageLayout',
    });
    expect(queryRibbonSearchIndex(items, 'page break preview')[0]).toMatchObject({
      commandId: 'viewPageBreakPreview',
      tab: 'view',
    });
    expect(queryRibbonSearchIndex(items, 'input message')[0]).toMatchObject({
      commandId: 'dataValidation',
      tab: 'data',
    });
    expect(queryRibbonSearchIndex(items, 'name manager')[0]).toMatchObject({
      commandId: 'namedRanges',
      tab: 'formulas',
    });
    expect(queryRibbonSearchIndex(items, 'edit links')[0]).toMatchObject({
      commandId: 'linksData',
      tab: 'data',
    });
    expect(queryRibbonSearchIndex(items, 'screen clipping')[0]).toMatchObject({
      commandId: 'screenshotInsert',
      tab: 'insert',
    });
    expect(queryRibbonSearchIndex(items, 'stock images')[0]).toMatchObject({
      commandId: 'pictureInsert',
      tab: 'insert',
    });
    expect(queryRibbonSearchIndex(items, 'recommended charts')[0]).toMatchObject({
      commandId: 'chartInsert',
      tab: 'insert',
    });
    expect(queryRibbonSearchIndex(items, 'manual calculation')[0]).toMatchObject({
      commandId: 'calcOptions',
      tab: 'formulas',
    });
    expect(queryRibbonSearchIndex(items, 'watch window')[0]).toMatchObject({
      commandId: 'watch',
      tab: 'formulas',
    });
    expect(queryRibbonSearchIndex(items, 'protect sheet')[0]).toMatchObject({
      commandId: 'protectReview',
      tab: 'review',
    });
  });

  it('uses shared static popularity boosts to break close textual matches', () => {
    const items = [
      {
        id: 'command:customBroad',
        kind: 'command' as const,
        label: 'Analyze',
        hint: 'pivot table helper',
        tab: 'insert' as const,
        commandId: 'customBroad',
        keywords: 'analyze pivot table helper',
      },
      {
        id: 'command:pivotTableInsert',
        kind: 'command' as const,
        label: 'Analyze',
        hint: 'pivot table helper',
        tab: 'insert' as const,
        commandId: 'pivotTableInsert',
        keywords: 'analyze pivot table helper',
      },
    ];

    expect(queryRibbonSearchIndex(items, 'pivot')[0]?.commandId).toBe('pivotTableInsert');
  });

  it('accepts host-provided usage priors without framework-specific search code', () => {
    const items = [
      {
        id: 'command:alpha',
        kind: 'command' as const,
        label: 'Analyze',
        hint: 'pivot table helper',
        tab: 'insert' as const,
        commandId: 'alpha',
        keywords: 'analyze pivot table helper',
      },
      {
        id: 'command:beta',
        kind: 'command' as const,
        label: 'Analyze',
        hint: 'pivot table helper',
        tab: 'insert' as const,
        commandId: 'beta',
        keywords: 'analyze pivot table helper',
      },
    ];

    expect(
      queryRibbonSearchIndex(items, 'pivot', 8, {
        usagePrior: { commandBoosts: { beta: 50 } },
      })[0]?.commandId,
    ).toBe('beta');
  });

  it('omits disabled commands unless explicitly requested', () => {
    const standard = buildRibbonSearchIndex('en');
    const withDisabled = buildRibbonSearchIndex('en', { includeDisabled: true });

    expect(standard.some((item) => item.commandId === 'helpSearch')).toBe(false);
    expect(withDisabled.find((item) => item.commandId === 'helpSearch')).toMatchObject({
      disabled: true,
      disabledReason: 'Coming soon',
      tab: 'help',
    });
    expect(queryRibbonSearchIndex(withDisabled, 'coming soon', 8)[0]).toMatchObject({
      commandId: 'helpSearch',
      disabled: true,
      disabledReason: 'Coming soon',
    });
  });

  it('keeps Search/Tell me disabled state aligned with activation metadata', () => {
    const items = buildRibbonSearchIndex('en', {
      includeDisabled: true,
      tabs: EXCEL365_STANDARD_RIBBON_TABS,
    });
    const activatableCommands = new Set(
      ribbonActivatableSurfaceCommandIds({ tabs: EXCEL365_STANDARD_RIBBON_TABS }),
    );
    const activationDisabledCommands = items
      .filter((item) => item.kind === 'command' && item.commandId)
      .filter((item) => activatableCommands.has(item.commandId ?? ''))
      .filter((item) => ribbonActivationForCommand(item.commandId ?? '').kind === 'disabled')
      .map((item) => item.commandId)
      .sort();

    expect(activationDisabledCommands).toEqual(['helpSearch']);
    expect(items.find((item) => item.commandId === 'helpSearch')).toMatchObject({
      disabled: true,
      disabledReason: 'Coming soon',
      tab: 'help',
    });
  });
});
