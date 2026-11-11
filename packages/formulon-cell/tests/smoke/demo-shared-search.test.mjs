import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it, vi } from 'vitest';
import { createCommandPalette } from '../../../../apps/demo-shared/command-palette.ts';
import {
  buildDemoPrintPreviewModel,
  buildDemoSearchItems,
  buildDemoSearchSuggestions,
  commitDemoBookName,
  createDemoStrings,
  createInitialDemoWorkbook,
  DEFAULT_DEMO_QUICK_ACCESS,
  DEMO_ICONS,
  DEMO_PRINTER_PROFILE_ID,
  DEMO_PRINTER_PROFILES,
  DEMO_QUICK_ACCESS_COMMANDS,
  DEMO_QUICK_ACCESS_GLYPHS,
  demoChangeLogEntry,
  demoQuickAccessEnabled,
  demoQuickAccessMenu,
  demoSelectionLabel,
  INITIAL_DEMO_SEARCH_STATE,
  isDemoFeatureOn,
  nextDemoFeatureOverrides,
  pushDemoChangeLog,
  queryDemoSearchItems,
  reduceDemoSearch,
  refreshDemoPrinterProfiles,
  resolveDemoSearchKey,
  saveDemoWorkbookToDownload,
  toggleDemoQuickAccess,
} from '../../../../apps/demo-shared/index.ts';
import { mount } from '../../src/index.ts';
import { createExcelRibbonSvg } from '../../src/toolbar/index.ts';
import { createHostElement, installDomStubs, uninstallDomStubs } from '../test-utils/dom.ts';

const repoRoot = join(import.meta.dirname, '../../../..');

const DEMO_HOST = {
  printerProfiles: DEMO_PRINTER_PROFILES,
  printerProfileId: DEMO_PRINTER_PROFILE_ID,
};

/** Mounts the demo's empty workbook with the demo host providers, runs `body`, then releases both. */
const withDemoInstance = async (body) => {
  installDomStubs();
  const workbook = await createInitialDemoWorkbook('?locale=en&fixture=empty');
  const { host, cleanup } = createHostElement();
  const instance = await mount(host, { workbook, host: DEMO_HOST });
  try {
    return await body(instance);
  } finally {
    instance.dispose();
    workbook.dispose();
    cleanup();
    uninstallDomStubs();
  }
};

const writeText = (instance, values) =>
  instance.applyChanges(
    Object.entries(values).map(([ref, input]) => ({ sheet: 0, at: ref, input })),
    { history: 'record' },
  );

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
    // The Mac title bar's redo slot repeats the last action and says so.
    expect(react.ja.quickAccess.buttons.redo).toBe('繰り返し');
    expect(react.en.quickAccess.buttons.redo).toBe('Repeat');
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

    const reactSource = readFileSync(join(repoRoot, 'apps/react-demo/src/DemoIcon.tsx'), 'utf8');
    const vueSource = readFileSync(join(repoRoot, 'apps/vue-demo/src/DemoIcon.vue'), 'utf8');

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
    const ui = createDemoStrings('React').en;
    const model = await withDemoInstance(async (instance) => {
      await instance.workbook
        .sheet(0)
        .updatePageMargins({ topIn: 0, rightIn: 0, bottomIn: 0, leftIn: 0 });
      return buildDemoPrintPreviewModel(ui, instance, 'Book1');
    });

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
    const ui = createDemoStrings('React').ja;
    const model = await withDemoInstance((instance) =>
      buildDemoPrintPreviewModel(ui, instance, 'ブック'),
    );

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

  it('drives demo Save download and status bar upload state through the shared helper', async () => {
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

    await withDemoInstance(async (instance) => {
      await saveDemoWorkbookToDownload({
        instance,
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
      expect(urlApi.createObjectURL.mock.calls[0][0].size).toBeGreaterThan(0);
      expect(urlApi.revokeObjectURL).toHaveBeenCalledWith('blob:demo');

      // An instance whose workbook cannot be written reports the failure in the status bar.
      const broken = new Proxy(instance, {
        get: (target, key) => {
          if (key === 'workbook') throw new Error('save failed');
          return Reflect.get(target, key);
        },
      });
      await saveDemoWorkbookToDownload({
        instance: broken,
        bookName: 'Broken',
        setUploadStatus: (status) => statuses.push(status),
        documentRef,
        urlApi,
        setTimeoutFn,
      });
    });

    expect(statuses.slice(-2)).toEqual(['saving', 'error']);
    expect(clicked).toHaveBeenCalledTimes(1);
  });

  it('toggles feature overrides relative to the preset default', () => {
    // `charts` is on by default in the full preset, so turning it off adds an override.
    const off = nextDemoFeatureOverrides({
      preset: 'full',
      overrides: {},
      features: {},
      id: 'charts',
    });
    expect(off).toEqual({ charts: false });
    expect(isDemoFeatureOn({ charts: false }, 'charts')).toBe(false);

    // Toggling back to the preset default drops the override.
    expect(
      nextDemoFeatureOverrides({
        preset: 'full',
        overrides: off,
        features: off,
        id: 'charts',
      }),
    ).toEqual({});
  });

  it('labels single cells and ranges for the selection readout', () => {
    expect(
      demoSelectionLabel({
        active: { row: 2, col: 1 },
        range: { r0: 2, c0: 1, r1: 2, c1: 1 },
      }),
    ).toBe('B3');
    expect(
      demoSelectionLabel({
        active: { row: 2, col: 1 },
        range: { r0: 2, c0: 1, r1: 4, c1: 3 },
      }),
    ).toBe('B3:D5 · 9');
  });

  it('maps search input keys to close, move and run actions', () => {
    expect(resolveDemoSearchKey('Escape', 1, 3)).toEqual({ kind: 'close' });
    expect(resolveDemoSearchKey('ArrowDown', -1, 3)).toEqual({ kind: 'move', index: 0 });
    expect(resolveDemoSearchKey('ArrowUp', -1, 3)).toEqual({ kind: 'move', index: 2 });
    expect(resolveDemoSearchKey('Enter', -1, 3)).toEqual({ kind: 'run', index: 0 });
    expect(resolveDemoSearchKey('Enter', 2, 3)).toEqual({ kind: 'run', index: 2 });
    expect(resolveDemoSearchKey('Enter', -1, 0)).toBeNull();
    expect(resolveDemoSearchKey('a', 0, 3)).toBeNull();
  });

  it('keeps the first ArrowDown move when it opens a closed search list', () => {
    const closed = { query: 'sum', open: false, activeIndex: -1 };
    const action = resolveDemoSearchKey('ArrowDown', closed.activeIndex, 3);
    expect(reduceDemoSearch(closed, { type: 'key', action })).toEqual({
      query: 'sum',
      open: true,
      activeIndex: 0,
    });
  });

  it('resets the active index on focus and typing, and reopens only while focused', () => {
    const moved = { query: 'a', open: true, activeIndex: 2 };
    expect(reduceDemoSearch(moved, { type: 'focus' })).toEqual({
      query: 'a',
      open: true,
      activeIndex: -1,
    });
    expect(reduceDemoSearch(moved, { type: 'input', value: 'ab', focused: true })).toEqual({
      query: 'ab',
      open: true,
      activeIndex: -1,
    });
    // Escape clears the input natively after blur; that change must not reopen the list.
    expect(reduceDemoSearch(moved, { type: 'input', value: '', focused: false })).toEqual({
      query: '',
      open: false,
      activeIndex: -1,
    });
  });

  it('closes on Escape, keeps the index on blur, tracks hover and resets after a run', () => {
    const open = { query: 'a', open: true, activeIndex: 1 };
    expect(reduceDemoSearch(open, { type: 'key', action: { kind: 'close' } })).toEqual({
      query: 'a',
      open: false,
      activeIndex: -1,
    });
    expect(reduceDemoSearch(open, { type: 'blur' })).toEqual({ ...open, open: false });
    expect(reduceDemoSearch(open, { type: 'hover', index: 2 }).activeIndex).toBe(2);
    expect(reduceDemoSearch(open, { type: 'reset' })).toEqual(INITIAL_DEMO_SEARCH_STATE);
  });

  it('keeps the change log newest-first and capped', async () => {
    await withDemoInstance((instance) => {
      writeText(instance, { A1: 'zero' });
      const event = (row) => ({
        sheet: 0,
        at: { row, col: 0 },
        value: instance.workbook.sheet(0).getValue(row, 0),
        origin: 'user',
      });
      expect(demoChangeLogEntry(instance.workbook, event(0))).toMatchObject({
        cell: 'A1',
        preview: '"zero"',
      });
      expect(demoChangeLogEntry(instance.workbook, event(1)).preview).toBe('∅');
      let log = [];
      for (let row = 0; row < 12; row += 1)
        log = pushDemoChangeLog(log, demoChangeLogEntry(instance.workbook, event(row)));
      expect(log).toHaveLength(8);
      expect(log[0]?.cell).toBe('A12');
    });
  });

  it('lists the Mac Quick Access customize menu in desktop order with the default checks', () => {
    const ui = createDemoStrings('React').ja;
    expect(
      demoQuickAccessMenu(ui, DEFAULT_DEMO_QUICK_ACCESS).map((i) => [i.label, i.checked]),
    ).toEqual([
      ['自動保存', true],
      ['新規作成', false],
      ['ホーム', true],
      ['保存', true],
      ['名前を付けて保存', true],
      ['ブックを送信', false],
      ['印刷', false],
      ['スペル チェック', false],
      ['元に戻す', true],
      ['繰り返し', true],
      ['昇順で並べ替え', false],
      ['降順で並べ替え', false],
    ]);
    expect(
      DEFAULT_DEMO_QUICK_ACCESS.map((id) => ui.quickAccess.buttons[id] ?? ui.quickAccess.autosave),
    ).toEqual(['自動保存', 'ホーム', '保存', '別名で保存', '元に戻す', '繰り返し']);
    expect(DEMO_QUICK_ACCESS_GLYPHS.redo).toBe('repeat');
    expect(DEMO_QUICK_ACCESS_GLYPHS.home).toBe('home');
  });

  it('keeps toggled Quick Access commands in customize-menu order', () => {
    const added = toggleDemoQuickAccess(DEFAULT_DEMO_QUICK_ACCESS, 'print');
    expect(added).toEqual(['autosave', 'home', 'save', 'saveAs', 'print', 'undo', 'redo']);
    expect(toggleDemoQuickAccess(added, 'home')).toEqual([
      'autosave',
      'save',
      'saveAs',
      'print',
      'undo',
      'redo',
    ]);
    expect(DEMO_QUICK_ACCESS_COMMANDS).toHaveLength(12);
  });

  it('follows undo and repeat history, and disables Send Workbook always', async () => {
    await withDemoInstance((instance) => {
      expect(demoQuickAccessEnabled('undo', instance)).toBe(false);
      expect(demoQuickAccessEnabled('redo', instance)).toBe(false);
      writeText(instance, { A1: 'typed' });
      expect(demoQuickAccessEnabled('undo', instance)).toBe(true);
      expect(demoQuickAccessEnabled('redo', instance)).toBe(instance.history.canRepeat);
      expect(demoQuickAccessEnabled('sendWorkbook', instance)).toBe(false);
      expect(demoQuickAccessEnabled('save', instance)).toBe(true);
    });
    expect(demoQuickAccessEnabled('undo', null)).toBe(false);
    expect(demoQuickAccessEnabled('redo', null)).toBe(false);
  });

  it('builds the Mac empty-query search popup sections', async () => {
    const ui = createDemoStrings('React').ja;
    const applyCommand = vi.fn(() => true);
    const { sections, opened } = await withDemoInstance((instance) => {
      const open = vi.spyOn(instance.dialogs, 'open').mockImplementation(() => {});
      const built = buildDemoSearchSuggestions(ui, {
        instance,
        setRibbonTab: vi.fn(),
        applyCommand,
      });
      built[1].items[0].run();
      built[1].items[1].run();
      return { sections: built, opened: open.mock.calls };
    });
    expect(
      sections.map((section) => [
        section.heading,
        section.items.map((item) => [item.label, item.disabled === true, item.submenu === true]),
      ]),
    ).toEqual([
      [
        'おすすめの操作',
        [
          ['ウインドウ枠の固定', false, true],
          ['ヘッダーとフッター...', false, false],
          ['シートの行を挿入', false, false],
          ['重複を削除', false, false],
          ['マクロ...', true, false],
        ],
      ],
      [
        '検索と置換',
        [
          ['検索する', false, false],
          ['検索と置換', false, false],
        ],
      ],
      [null, [['検索に関するフィードバック', true, false]]],
    ]);
    expect(ui.searchSuggestions.button).toBe('検索 (Cmd + Ctrl + U)');
    expect(opened).toEqual([
      ['findReplace', { tab: 'find' }],
      ['findReplace', { tab: 'replace' }],
    ]);
    sections[0].items[2].run();
    expect(applyCommand).toHaveBeenCalledWith('insertRows');
    expect(sections[2].items[0].glyph).toBe('feedback');
  });

  it('draws every Quick Access and search glyph from the core icon kit', () => {
    const ui = createDemoStrings('React').ja;
    const deps = { instance: null, setRibbonTab: () => {}, applyCommand: () => false };
    const names = [
      ...Object.values(DEMO_QUICK_ACCESS_GLYPHS).filter(Boolean),
      ...buildDemoSearchSuggestions(ui, deps).flatMap((s) => s.items.map((i) => i.glyph)),
      'find',
      'moreCommands',
    ];
    expect(names.filter((name) => createExcelRibbonSvg(name, 'x') === null)).toEqual([]);
  });

  it('commits a document-menu rename as the bare book name', () => {
    expect(commitDemoBookName('  Budget.xlsx ', 'Book1')).toBe('Budget');
    expect(commitDemoBookName('Q3 report.XLSM', 'Book1')).toBe('Q3 report');
    expect(commitDemoBookName('   ', 'Book1')).toBe('Book1');
    expect(commitDemoBookName('.xlsx', 'Book1')).toBe('Book1');
    const ui = createDemoStrings('React').ja;
    expect(ui.documentMenu).toEqual({
      button: 'ドキュメントの操作',
      name: '名前:',
      location: '場所:',
      versionHistory: 'バージョン履歴の表示',
    });
    expect(ui.comments).toBe('コメント');
    expect(
      ['comments', 'folder', 'versionHistory', 'share'].filter(
        (name) => createExcelRibbonSvg(name, 'x') === null,
      ),
    ).toEqual([]);
  });
});
