import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it, vi } from 'vitest';
import { createCommandPalette } from '../../../../apps/demo-shared/command-palette.ts';
import {
  buildDemoPrintPreviewModel,
  buildDemoSearchItems,
  createDemoStrings,
  DEMO_ICONS,
  DEMO_PRINTER_PROFILE_ID,
  DEMO_PRINTER_PROFILES,
  demoCommandText,
  queryDemoSearchItems,
  refreshDemoPrinterProfiles,
  saveDemoWorkbookToDownload,
} from '../../../../apps/demo-shared/index.ts';
import { createSpreadsheetStore, mutators, WorkbookHandle } from '../../src/index.ts';

const repoRoot = join(import.meta.dirname, '../../../..');

describe('demo-shared Search/Tell me items', () => {
  it('keeps visible workbook chrome framework-neutral for the Excel baseline', () => {
    const react = createDemoStrings('React');
    const vue = createDemoStrings('Vue');

    expect(react.en.workbook).toBe('Workbook');
    expect(vue.en.workbook).toBe('Workbook');
    expect(react.ja.workbook).toBe('ブック');
    expect(vue.ja.workbook).toBe('ブック');
    expect(react.en.quickAccessToolbar).toBe('Quick Access Toolbar');
    expect(react.ja.quickAccessToolbar).toBe('クイック アクセス ツール バー');
    expect(react.en.optionsPanel).toBe('Options panel');
    expect(react.ja.optionsPanel).toBe('オプション パネル');
    expect(react.en.undo).toBe('Undo');
    expect(react.en.redo).toBe('Redo');
    expect(react.ja.undo).toBe('元に戻す');
    expect(react.ja.redo).toBe('やり直し');
    expect(react.en.demoChrome).toBe('Demo chrome');
    expect(react.ja.demoChrome).toBe('デモ表示');
    expect(react.en.themeLabels.paper).toBe('Light');
    expect(react.ja.themeLabels.paper).toBe('ライト');
    expect(react.en.presets.full.label).toBe('Full');
    expect(react.ja.presets.full.label).toBe('フル');
    expect(react.en.featureGroupLabels.Editing).toBe('Editing');
    expect(react.ja.featureGroupLabels.Editing).toBe('編集');
    expect(react.en.featureLabels.formulaBar).toBe('Formula bar');
    expect(react.ja.featureLabels.formulaBar).toBe('数式バー');
    expect(react.en.featureLabels.pasteSpecial).toBe('Paste special');
    expect(react.ja.featureLabels.pasteSpecial).toBe('形式を選択して貼り付け');
    expect(react.en.spreadsheetRibbon).toBe('Spreadsheet ribbon');
    expect(react.ja.spreadsheetRibbon).toBe('スプレッドシート リボン');
    expect(react.en.cellChangeLog).toBe('Cell change log');
    expect(react.ja.cellChangeLog).toBe('セル変更ログ');
    expect(react.en.noIssuesFound).toBe('No issues found.');
    expect(react.ja.noIssuesFound).toBe('問題は見つかりませんでした。');
    expect(react.en.loadingEngine).toBe('Loading engine...');
    expect(react.ja.loadingEngine).toBe('エンジンを読み込んでいます...');
    expect(react.en.command).toBe('Command');
    expect(react.ja.command).toBe('コマンド');
    expect(demoCommandText('en').spellingReview).toBe('Spelling Review');
    expect(demoCommandText('ja').spellingReview).toBe('スペル チェック');
    expect(react.en.cancel).toBe('Cancel');
    expect(react.ja.cancel).toBe('キャンセル');
    expect(react.en.run).toBe('Run');
    expect(react.ja.run).toBe('実行');
    expect(react.en.backstageSub).toBe('Workbook · spreadsheet layout');
    expect(vue.en.backstageSub).toBe('Workbook · spreadsheet layout');
    expect(react.ja.backstageSub).toBe('ブック · スプレッドシート レイアウト');
    expect(vue.ja.backstageSub).toBe('ブック · スプレッドシート レイアウト');
    expect(`${react.en.workbook} ${react.en.backstageSub}`).not.toMatch(/React|Vue/);
    expect(`${react.ja.workbook} ${react.ja.backstageSub}`).not.toMatch(/React|Vue/);
  });

  it('keeps demo Quick Access and search icons on semantic colored SVG segments', () => {
    for (const iconName of ['app', 'save', 'undo', 'redo', 'search']) {
      const segments = DEMO_ICONS[iconName];

      expect(segments.length, iconName).toBeGreaterThan(1);
      expect(segments.some((segment) => segment.stroke && segment.stroke !== 'currentColor')).toBe(
        true,
      );
    }

    expect(DEMO_ICONS.app.some((segment) => segment.fill === '#107c41')).toBe(true);
    expect(DEMO_ICONS.save.some((segment) => segment.fill === '#2f75b5')).toBe(true);
    expect(DEMO_ICONS.search.some((segment) => segment.stroke === '#107c41')).toBe(true);

    const reactSource = readFileSync(join(repoRoot, 'apps/react-demo/src/App.tsx'), 'utf8');
    const vueSource = readFileSync(join(repoRoot, 'apps/vue-demo/src/App.vue'), 'utf8');

    expect(reactSource).toContain("fill={segment.fill ?? 'none'}");
    expect(reactSource).toContain("stroke={segment.stroke ?? 'currentColor'}");
    expect(vueSource).toContain(':fill="segment.fill ?? \'none\'"');
    expect(vueSource).toContain(':stroke="segment.stroke ?? \'currentColor\'"');
  });

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
    const disabled = container.querySelector('.fc-tb__command-item');
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
    expect(model.settings).toContainEqual({
      label: 'Orientation',
      value: 'Portrait Orientation',
    });
    expect(model.previewHtml).toContain(
      '@page { size: A4 portrait; margin: 0.16in 0.14in 0.18in 0.14in; }',
    );
  });

  it('localizes the Backstage print preview model for Japanese Excel chrome', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const ui = createDemoStrings('React').ja;

    const model = buildDemoPrintPreviewModel(ui, { store, workbook }, 'ブック');

    expect(model.title).toBe('印刷');
    expect(model.subtitle).toBe('ブック');
    expect(model.printLabel).toBe('印刷');
    expect(model.pdfLabel).toBe('PDF にエクスポート');
    expect(model.pageSetupLabel).toBe('ページ設定');
    expect(model.previewTitle).toBe('ページ 1');
    expect(model.previewHint).toBe('プレビューはアクティブ シートのページ設定を反映します。');
    expect(model.settings).toEqual(
      expect.arrayContaining([
        { label: 'アクティブ シート', value: '1' },
        { label: '印刷の向き', value: '縦方向' },
        { label: '用紙サイズ', value: 'A4' },
        { label: 'プリンター', value: 'Demo Office Printer - A4' },
        { label: '余白', value: '0.75" / 0.7" / 0.75" / 0.7"' },
        { label: '拡大縮小', value: '100%' },
        { label: '印刷範囲', value: '印刷範囲なし' },
      ]),
    );
    expect(
      [
        model.title,
        model.subtitle,
        model.printLabel,
        model.pdfLabel,
        model.pageSetupLabel,
        model.previewTitle,
        model.previewHint,
        ...model.settings.flatMap((setting) => [setting.label, setting.value]),
      ].join('\n'),
    ).not.toMatch(/React|Vue|portrait|landscape|Orientation|Print area|Active sheet/);
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
