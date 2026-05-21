import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it, vi } from 'vitest';
import { createCommandPalette } from '../../../../apps/demo-shared/command-palette.ts';
import {
  buildDemoPrintPreviewModel,
  buildDemoSearchItems,
  createDemoStrings,
  DEMO_PRINTER_PROFILE_ID,
  DEMO_PRINTER_PROFILES,
  queryDemoSearchItems,
  refreshDemoPrinterProfiles,
  saveDemoWorkbookToDownload,
} from '../../../../apps/demo-shared/index.ts';
import { createSpreadsheetStore, mutators, WorkbookHandle } from '../../src/index.ts';

const repoRoot = join(import.meta.dirname, '../../../..');

describe('demo-shared Search/Tell me items', () => {
  it('keeps disabled ribbon command reasons visible through the shared demo search model', () => {
    const setRibbonTab = vi.fn();
    const applyRibbonCommand = vi.fn(() => false);
    const items = buildDemoSearchItems([], 'en', setRibbonTab, applyRibbonCommand);

    const helpSearch = queryDemoSearchItems(items, 'coming soon', 8).find(
      (item) => item.commandId === 'helpSearch',
    );

    expect(helpSearch).toMatchObject({
      commandId: 'helpSearch',
      disabled: true,
      disabledReason: 'Coming soon',
      tab: 'help',
    });
    expect(helpSearch?.hint).toContain('Coming soon');

    helpSearch?.run();
    expect(applyRibbonCommand).toHaveBeenCalledWith('helpSearch');
    expect(setRibbonTab).toHaveBeenCalledWith('help');
  });

  it('uses the same disabled hint path for Japanese demo search results', () => {
    const items = buildDemoSearchItems(
      [],
      'ja',
      vi.fn(),
      vi.fn(() => false),
    );

    const helpSearch = queryDemoSearchItems(items, '未実装', 8).find(
      (item) => item.commandId === 'helpSearch',
    );

    expect(helpSearch?.disabled).toBe(true);
    expect(helpSearch?.disabledReason).toBe('未実装');
    expect(helpSearch?.hint).toContain('未実装');
  });

  it('keeps standalone command palette disabled results aligned with shared search', () => {
    const container = document.createElement('div');
    const input = document.createElement('input');
    document.body.append(container, input);
    const applyCommand = vi.fn(() => false);
    const palette = createCommandPalette({
      input,
      container,
      ribbonLang: 'en',
      applyCommand,
      selectTab: vi.fn(),
    });

    input.value = 'coming soon';
    input.dispatchEvent(new Event('input', { bubbles: true }));
    const disabled = container.querySelector('.demo__command-item');
    expect(disabled?.textContent).toContain('Coming soon');
    expect(disabled?.getAttribute('aria-disabled')).toBe('true');
    expect(disabled?.dataset.disabledReason).toBe('Coming soon');
    disabled?.click();
    expect(applyCommand).toHaveBeenCalledWith('helpSearch');

    palette.dispose();
    container.remove();
    input.remove();
  });

  it('keeps standalone command palette on the shared disabled search path', () => {
    const source = readFileSync(join(repoRoot, 'apps/demo-shared/command-palette.ts'), 'utf8');

    expect(source).toContain('buildRibbonSearchIndex(ribbonLang, { includeDisabled: true })');
    expect(source).toContain('projectDisabledReason(btn, cmd.disabledReason');
    expect(source).not.toContain('ribbonText:');
  });

  it('preserves shared ribbon aliases through the React/Vue demo search model', () => {
    const setRibbonTab = vi.fn();
    const applyRibbonCommand = vi.fn(() => false);
    const items = buildDemoSearchItems([], 'en', setRibbonTab, applyRibbonCommand);

    const lockPanes = queryDemoSearchItems(items, 'lock panes', 8)[0];
    expect(lockPanes).toMatchObject({
      commandId: 'freeze',
      tab: 'view',
    });
    lockPanes?.run();
    expect(applyRibbonCommand).toHaveBeenCalledWith('freeze');

    expect(queryDemoSearchItems(items, 'screen clipping', 8)[0]).toMatchObject({
      commandId: 'screenshotInsert',
      tab: 'insert',
    });
    expect(queryDemoSearchItems(items, 'name manager', 8)[0]).toMatchObject({
      commandId: 'namedRanges',
      tab: 'formulas',
    });
    expect(queryDemoSearchItems(items, 'stock images', 8)[0]).toMatchObject({
      commandId: 'pictureInsert',
      tab: 'insert',
    });
    expect(queryDemoSearchItems(items, 'recommended charts', 8)[0]).toMatchObject({
      commandId: 'chartInsert',
      tab: 'insert',
    });
    expect(queryDemoSearchItems(items, 'external links', 8)[0]).toMatchObject({
      commandId: 'linksData',
      tab: 'data',
    });
    expect(queryDemoSearchItems(items, 'protect sheet password', 8)[0]).toMatchObject({
      commandId: 'protectReview',
      tab: 'review',
    });
    expect(queryDemoSearchItems(items, 'bring to front', 8)[0]).toMatchObject({
      commandId: 'arrangeObjectsPageLayout',
      tab: 'pageLayout',
    });
    expect(queryDemoSearchItems(items, 'check accessibility', 8)[0]).toMatchObject({
      commandId: 'accessibility',
      tab: 'review',
    });
    expect(queryDemoSearchItems(items, 'combine cells', 8)[0]).toMatchObject({
      commandId: 'merge',
      tab: 'home',
    });
    expect(queryDemoSearchItems(items, 'show formula bar', 8)[0]).toMatchObject({
      commandId: 'viewFormulaBar',
      tab: 'view',
    });
    expect(queryDemoSearchItems(items, 'page break preview', 8)[0]).toMatchObject({
      commandId: 'viewPageBreakPreview',
      tab: 'view',
    });
  });

  it('keeps demo search aligned with the visible standard ribbon tab surface', () => {
    const setRibbonTab = vi.fn();
    const applyRibbonCommand = vi.fn(() => false);
    const standardItems = buildDemoSearchItems([], 'en', setRibbonTab, applyRibbonCommand);

    expect(queryDemoSearchItems(standardItems, 'script', 8)).not.toEqual(
      expect.arrayContaining([expect.objectContaining({ commandId: 'script' })]),
    );
    expect(queryDemoSearchItems(standardItems, 'pdf', 8)).not.toEqual(
      expect.arrayContaining([expect.objectContaining({ commandId: 'pdf' })]),
    );

    const optionalItems = buildDemoSearchItems([], 'en', setRibbonTab, applyRibbonCommand, [
      'automate',
      'acrobat',
    ]);
    expect(queryDemoSearchItems(optionalItems, 'script', 8)[0]).toMatchObject({
      commandId: 'script',
      tab: 'automate',
    });
    expect(queryDemoSearchItems(optionalItems, 'pdf', 8)[0]).toMatchObject({
      commandId: 'pdf',
      tab: 'acrobat',
    });
  });

  it('exposes a shared host-printer stub for React and Vue demos', async () => {
    expect(DEMO_PRINTER_PROFILE_ID).toBe('demo-office-a4');
    expect(DEMO_PRINTER_PROFILES).toContainEqual({
      id: 'demo-office-a4',
      name: 'Demo Office Printer - A4',
      paperSize: 'A4',
      orientation: 'portrait',
      printableBounds: { top: 0.16, right: 0.14, bottom: 0.18, left: 0.14 },
    });

    const refreshed = await refreshDemoPrinterProfiles();
    expect(refreshed).toEqual(DEMO_PRINTER_PROFILES);
    expect(refreshed).not.toBe(DEMO_PRINTER_PROFILES);
    expect(refreshed[0]?.printableBounds).not.toBe(DEMO_PRINTER_PROFILES[0]?.printableBounds);
  });

  it('feeds the shared host-printer stub into the Backstage print preview model', async () => {
    const store = createSpreadsheetStore();
    mutators.setPageSetup(store, 0, {
      margins: { top: 0, right: 0, bottom: 0, left: 0 },
    });
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const ui = createDemoStrings('React').en;

    const model = buildDemoPrintPreviewModel(ui, { store, workbook }, 'Book1');

    expect(model.settings).toContainEqual({
      label: 'Printer',
      value: 'Demo Office Printer - A4',
    });
    expect(model.settings).toContainEqual({
      label: 'Minimum margins',
      value: '0.16" / 0.14" / 0.18" / 0.14"',
    });
    expect(model.previewHtml).toContain(
      '@page { size: A4 portrait; margin: 0.16in 0.14in 0.18in 0.14in; }',
    );
  });

  it('drives demo Save download and status bar upload state through the shared helper', () => {
    const statuses = [];
    const clicked = vi.fn();
    const appended = [];
    const removed = [];
    const anchor = { href: '', download: '', click: clicked };
    const documentRef = {
      body: {
        appendChild: (el) => appended.push(el),
        removeChild: (el) => removed.push(el),
      },
      createElement: (tag) => {
        expect(tag).toBe('a');
        return anchor;
      },
    };
    const urlApi = {
      createObjectURL: vi.fn(() => 'blob:demo'),
      revokeObjectURL: vi.fn(),
    };
    const setTimeoutFn = (handler) => {
      handler();
      return 1;
    };

    saveDemoWorkbookToDownload({
      instance: { workbook: { save: () => new Uint8Array([1, 2, 3]) } },
      bookName: 'Book1',
      setUploadStatus: (status) => statuses.push(status),
      documentRef,
      urlApi,
      setTimeoutFn,
    });

    expect(statuses).toEqual(['saving', 'saved']);
    expect(anchor).toMatchObject({ href: 'blob:demo', download: 'Book1.xlsx' });
    expect(clicked).toHaveBeenCalledTimes(1);
    expect(appended).toEqual([anchor]);
    expect(removed).toEqual([anchor]);
    expect(urlApi.revokeObjectURL).toHaveBeenCalledWith('blob:demo');

    saveDemoWorkbookToDownload({
      instance: {
        workbook: {
          save: () => {
            throw new Error('save failed');
          },
        },
      },
      bookName: 'Broken',
      setUploadStatus: (status) => statuses.push(status),
      documentRef,
      urlApi,
      setTimeoutFn,
    });

    expect(statuses.slice(-2)).toEqual(['saving', 'error']);
  });
});
