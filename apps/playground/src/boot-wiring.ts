import {
  attachFilterDropdown,
  type FeatureFlags,
  recordFormatChange,
  Spreadsheet,
  type SpreadsheetInstance,
  WorkbookHandle,
} from '@libraz/formulon-cell';

import { applyFixture, isFixtureName } from './fixtures.js';

// Theme name as exposed by `@libraz/formulon-cell`. Kept inline to avoid a
// dependency on the playground-only `CoreTheme` alias declared in main.ts.
type CoreTheme = 'paper' | 'ink' | 'contrast';
type UiTheme = 'light' | 'dark' | 'contrast';

type FilterDropdown = ReturnType<typeof attachFilterDropdown>;

type ShellTextLike = {
  readonly saved: string;
  readonly ready: string;
};

export interface BootWiringCtx {
  // Spreadsheet instance plumbing.
  getInst: () => SpreadsheetInstance | null;
  setInst: (next: SpreadsheetInstance | null) => void;

  // Locale / theme.
  ribbonLang: 'ja' | 'en';
  localeParam: string | null;
  getUiTheme: () => UiTheme;
  toCore: (theme: UiTheme) => CoreTheme;

  // DOM anchors.
  sheetEl: HTMLElement;
  enginePill: HTMLElement | null;
  statusEngine: HTMLElement | null;
  docState: HTMLElement | null;
  statusState: HTMLElement | null;

  // Status / format projection hooks (late-bound through thunks where the
  // underlying values are assigned after this factory runs).
  projectStatus: () => void;
  projectFormatToolbar: () => void;
  refreshObjectsBadge: (
    source: 'passthroughs' | 'tables',
    detail: { count: number; byCategory?: Record<string, number> },
  ) => void;
  markDirty: () => void;
  refreshZoom: () => void;
  renderSheetTabs: () => void;

  // Misc shell.
  shellText: ShellTextLike;
  bootParams: URLSearchParams;
  seed: (wb: WorkbookHandle) => void;
  playgroundFeatureFlags: () => FeatureFlags;

  // Filter-dropdown plumbing (boot owns the lifecycle).
  setFilterDropdown: (drop: FilterDropdown | null) => void;

  // CSS class toggled on the active toolbar button.
  activeClass: string;
}

export interface BootWiringApi {
  boot: () => Promise<void>;
  openCommentDialog: () => void;
  wireFormat: (
    id: string,
    fn: (
      state: ReturnType<SpreadsheetInstance['store']['getState']>,
      store: SpreadsheetInstance['store'],
    ) => void,
  ) => void;
  formatPainterBtn: HTMLElement | null;
}

export const createBootWiring = (ctx: BootWiringCtx): BootWiringApi => {
  const {
    getInst,
    setInst,
    ribbonLang: _ribbonLang,
    localeParam,
    getUiTheme,
    toCore,
    sheetEl,
    enginePill,
    statusEngine,
    docState,
    statusState,
    projectStatus,
    projectFormatToolbar,
    refreshObjectsBadge,
    markDirty,
    refreshZoom,
    renderSheetTabs,
    shellText,
    bootParams,
    seed,
    playgroundFeatureFlags,
    setFilterDropdown,
    activeClass,
  } = ctx;
  // `ribbonLang` is currently unused by the extracted block but kept on the
  // context for symmetry with sibling wirings (and so future tweaks don't have
  // to thread it back through).
  void _ribbonLang;

  // ── Format Painter (module-scope DOM ref + sticky timer) ───────────────
  const formatPainterBtn = document.getElementById('btn-format-painter');
  let painterStickyTimer: number | null = null;
  formatPainterBtn?.addEventListener('click', () => {
    const inst = getInst();
    if (!inst) return;
    // Defer one-shot activation briefly so a follow-up click within the
    // dblclick window can promote it to sticky without painting twice.
    if (painterStickyTimer != null) return;
    painterStickyTimer = window.setTimeout(() => {
      painterStickyTimer = null;
      const i = getInst();
      if (!i) return;
      const fp = i.formatPainter;
      if (!fp) return;
      if (fp.isActive()) fp.deactivate();
      else fp.activate(false);
      sheetEl.focus();
      formatPainterBtn?.classList.toggle(activeClass, fp.isActive());
    }, 220);
  });
  formatPainterBtn?.addEventListener('dblclick', () => {
    const inst = getInst();
    if (!inst) return;
    if (painterStickyTimer != null) {
      clearTimeout(painterStickyTimer);
      painterStickyTimer = null;
    }
    const fp = inst.formatPainter;
    if (!fp) return;
    fp.activate(true);
    sheetEl.focus();
    formatPainterBtn?.classList.toggle(activeClass, fp.isActive());
  });

  // ── Comment dialog opener (shared by Insert ribbon + Review ribbon) ────
  const openCommentDialog = (): void => {
    getInst()?.openCommentDialog();
  };

  // ── Toolbar format wiring helper (Bold / Italic / Align / …) ───────────
  const wireFormat = (
    id: string,
    fn: (
      state: ReturnType<SpreadsheetInstance['store']['getState']>,
      store: SpreadsheetInstance['store'],
    ) => void,
  ): void => {
    document.getElementById(id)?.addEventListener('click', () => {
      const i = getInst();
      if (!i) return;
      // Wrap each toolbar mutation so Cmd+Z reverts the format change.
      recordFormatChange(i.history, i.store, () => {
        fn(i.store.getState(), i.store);
      });
      sheetEl.focus();
    });
  };

  // ── boot() ─────────────────────────────────────────────────────────────
  const boot = async (): Promise<void> => {
    // Default to the real WASM engine. Pass ?engine=stub to force the JS stub
    // for explicit demos or behavior diffs.
    const params = new URLSearchParams(window.location.search);
    const preferStub = params.get('engine') === 'stub';
    const wb = await WorkbookHandle.createDefault({
      preferStub,
      onFallback: (reason) => {
        // eslint-disable-next-line no-console
        console.info('[formulon-cell]', reason);
      },
    });
    // mount.ts only runs `seed` on workbooks it owns. We construct `wb` here so
    // we can read `isStub` / `version` for the engine pill before mounting,
    // which means we have to seed the workbook ourselves. `?fixture=empty`
    // (used by E2E specs that need a deterministic blank workbook) skips this.
    if (bootParams.get('fixture') !== 'empty') {
      seed(wb);
    }

    const inst = await Spreadsheet.mount(sheetEl, {
      theme: toCore(getUiTheme()),
      workbook: wb,
      locale: localeParam === 'en' ? 'en' : 'ja',
      features: playgroundFeatureFlags(),
    });
    setInst(inst);
    // Debug-only: expose for browser console / e2e poking. Safe to leave on the
    // playground build; the core package never references this global.
    (window as unknown as { __fcInst?: SpreadsheetInstance }).__fcInst = inst;

    // Visual-regression fixtures. `?fixture=cf|sparkline|selection|frozen`
    // replaces the default seed with a deterministic shape.
    const fixtureParam = bootParams.get('fixture');
    if (fixtureParam && isFixtureName(fixtureParam)) {
      applyFixture(fixtureParam, wb, inst);
    }

    setFilterDropdown(attachFilterDropdown({ store: inst.store }));

    // Read-only badge — chart/drawing/pivot counts and spreadsheet Tables. Hidden
    //  until the loaded workbook actually carries any of these objects.
    inst.host.addEventListener('fc:passthroughs', (ev) => {
      const e = ev as CustomEvent<{ count: number; byCategory: Record<string, number> }>;
      refreshObjectsBadge('passthroughs', e.detail);
    });
    inst.host.addEventListener('fc:tables', (ev) => {
      const e = ev as CustomEvent<{ count: number }>;
      refreshObjectsBadge('tables', e.detail);
    });
    // Header chevron click → mount.ts owns the `fc:openfilter` listener and
    // opens its own dropdown. The playground keeps its `filterDropdown` only
    // for the sort menu's "filter" action.

    const engineLabel = wb.isStub ? 'stub engine' : `formulon ${wb.version}`;
    if (enginePill) enginePill.textContent = `engine · ${engineLabel}`;
    if (statusEngine) statusEngine.textContent = engineLabel;
    if (docState) docState.textContent = shellText.saved;
    if (statusState) statusState.textContent = shellText.ready;

    inst.store.subscribe(() => {
      projectStatus();
      projectFormatToolbar();
      markDirty();
      refreshZoom();
    });
    projectStatus();
    projectFormatToolbar();
    renderSheetTabs();
    refreshZoom();

    // Reflect Format Painter state on the toolbar button (any path can deactivate
    // it — Esc, post-paint, or programmatic).
    inst.formatPainter?.subscribe((active, sticky) => {
      formatPainterBtn?.classList.toggle(activeClass, active);
      formatPainterBtn?.classList.toggle('app__tool--sticky', active && sticky);
    });
  };

  return {
    boot,
    openCommentDialog,
    wireFormat,
    formatPainterBtn,
  };
};
