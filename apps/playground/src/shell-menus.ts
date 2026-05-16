import {
  applyMerge,
  applyUnmerge,
  isWorkbookStructureProtected,
  mutators,
  recordFormatChange,
  type SpreadsheetInstance,
  toggleWrap,
  WorkbookHandle,
} from '@libraz/formulon-cell';

export interface ShellMenusCtx {
  getInst: () => SpreadsheetInstance | null;
  ribbonLang: 'ja' | 'en';
  // DOM hooks
  ribbonRoot: HTMLElement | null;
  sheetEl: HTMLElement;
  docState: HTMLElement | null;
  // Shell i18n
  shellText: { saved: string; ready: string } & Record<string, string>;
  // xlsx-io hooks
  openFileMenu: () => void;
  closeFileMenu: () => void;
  triggerOpen: () => void;
  triggerSave: () => void;
  triggerSaveAs: () => Promise<void>;
  loadXlsxFile: (file: File) => Promise<void>;
  inspectWorkbookFromBackstage: () => void;
  setDocName: (name: string) => void;
  // Other host hooks
  renderSheetTabs: () => void;
  applyUiTheme: (theme: 'light' | 'dark' | 'contrast') => void;
  getUiTheme: () => 'light' | 'dark' | 'contrast';
  applyProtectAction: (
    action: 'protect-workbook' | 'unprotect-workbook' | 'protect-sheet' | 'unprotect-sheet',
  ) => Promise<void>;
  closeBackstage: (focusTab?: boolean) => void;
}

export interface ShellMenusApi {
  refreshViewMenu: () => void;
}

export const createShellMenus = (ctx: ShellMenusCtx): ShellMenusApi => {
  const {
    getInst,
    ribbonRoot,
    sheetEl,
    docState,
    shellText,
    openFileMenu,
    closeFileMenu,
    triggerOpen,
    triggerSave,
    triggerSaveAs,
    loadXlsxFile,
    inspectWorkbookFromBackstage,
    setDocName,
    renderSheetTabs,
    applyUiTheme,
    getUiTheme,
    applyProtectAction,
    closeBackstage,
  } = ctx;

  const themeToggle = document.getElementById('theme-toggle') as HTMLButtonElement | null;

  themeToggle?.addEventListener('click', () => {
    applyUiTheme(getUiTheme() === 'dark' ? 'light' : 'dark');
  });

  // ── File menu (New / Open / Save / Save As) ───────────────────────────────
  const fileMenuBtn = document.getElementById('menu-file');
  const fileMenuDrop = document.getElementById('menu-file-dropdown');
  const fileInput = document.getElementById('file-input') as HTMLInputElement | null;

  fileMenuBtn?.addEventListener('click', (e) => {
    e.stopPropagation();
    if (!fileMenuDrop) return;
    if (fileMenuDrop.hidden) openFileMenu();
    else closeFileMenu();
  });

  document.addEventListener('mousedown', (e) => {
    if (!fileMenuDrop || fileMenuDrop.hidden) return;
    if (fileMenuDrop.contains(e.target as Node)) return;
    if (fileMenuBtn?.contains(e.target as Node)) return;
    closeFileMenu();
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && !fileMenuDrop?.hidden) closeFileMenu();
  });

  fileInput?.addEventListener('change', () => {
    const f = fileInput.files?.[0];
    if (f) void loadXlsxFile(f);
    fileInput.value = ''; // allow same-file re-open
  });

  fileMenuDrop?.querySelectorAll<HTMLButtonElement>('[data-file]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const action = btn.dataset.file;
      closeFileMenu();
      const inst = getInst();
      if (!inst) return;
      if (action === 'new') {
        void (async () => {
          const next = await WorkbookHandle.createDefault();
          await getInst()?.setWorkbook(next);
          setDocName('Book1');
          if (docState) docState.textContent = shellText.saved;
          renderSheetTabs();
        })();
      } else if (action === 'open') {
        triggerOpen();
      } else if (action === 'save') {
        triggerSave();
      } else if (action === 'save-as') {
        void triggerSaveAs();
      }
    });
  });

  ribbonRoot?.addEventListener('click', (event) => {
    const button = (event.target as Element | null)?.closest<HTMLButtonElement>(
      '[data-backstage-action]',
    );
    if (!button || button.disabled) return;
    const action = button.dataset.backstageAction;
    if (!action || action === 'info') return;
    event.preventDefault();
    event.stopPropagation();
    const inst = getInst();
    if (action === 'back') {
      closeBackstage(true);
    } else if (action === 'new') {
      closeBackstage();
      void (async () => {
        const next = await WorkbookHandle.createDefault();
        await getInst()?.setWorkbook(next);
        setDocName('Book1');
        if (docState) docState.textContent = shellText.saved;
        renderSheetTabs();
      })();
    } else if (action === 'open') {
      closeBackstage();
      triggerOpen();
    } else if (action === 'save') {
      closeBackstage();
      triggerSave();
    } else if (action === 'save-as') {
      closeBackstage();
      void triggerSaveAs();
    } else if (action === 'print') {
      closeBackstage();
      inst?.print('print');
    } else if (action === 'options') {
      closeBackstage();
      inst?.openIterativeDialog();
    } else if (action === 'protect-workbook') {
      closeBackstage();
      void applyProtectAction(
        inst && isWorkbookStructureProtected(inst.store.getState())
          ? 'unprotect-workbook'
          : 'protect-workbook',
      );
    } else if (action === 'inspect-workbook') {
      closeBackstage();
      inspectWorkbookFromBackstage();
    } else if (action === 'links') {
      closeBackstage();
      inst?.openExternalLinksDialog();
    }
  });

  // Drag & drop xlsx onto the page.
  window.addEventListener('dragover', (e) => {
    if (!e.dataTransfer) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
  });
  window.addEventListener('drop', (e) => {
    e.preventDefault();
    const f = e.dataTransfer?.files?.[0];
    if (!f) return;
    if (!/\.xlsx?$/i.test(f.name)) return;
    void loadXlsxFile(f);
  });

  // Ctrl/Cmd-O / Ctrl/Cmd-S / Ctrl/Cmd-N for file actions.
  window.addEventListener('keydown', (e) => {
    if (!(e.ctrlKey || e.metaKey)) return;
    const k = e.key.toLowerCase();
    if (k === 'o') {
      e.preventDefault();
      triggerOpen();
    } else if (k === 's') {
      e.preventDefault();
      if (e.shiftKey) void triggerSaveAs();
      else triggerSave();
    } else if (k === 'n' && !e.shiftKey) {
      // Ctrl+N — create a fresh workbook in place.
      e.preventDefault();
      void (async () => {
        const next = await WorkbookHandle.createDefault();
        await getInst()?.setWorkbook(next);
        setDocName('Book1');
        renderSheetTabs();
      })();
    }
  });

  // ── View menu (Show Formulas / R1C1 / Grid / Headers toggles) ────────────
  const viewBtn = document.getElementById('menu-view');
  const viewDrop = document.getElementById('menu-view-dropdown');
  const closeViewMenu = (): void => {
    if (!viewDrop) return;
    viewDrop.hidden = true;
    viewBtn?.setAttribute('aria-expanded', 'false');
  };
  const refreshViewMenu = (): void => {
    const inst = getInst();
    if (!inst || !viewDrop) return;
    const ui = inst.store.getState().ui;
    const update = (action: string, on: boolean): void => {
      const item = viewDrop.querySelector<HTMLElement>(`[data-view="${action}"] [data-fc-check]`);
      if (item) item.textContent = on ? '✓' : '';
    };
    update('show-formulas', !!ui.showFormulas);
    update('r1c1', !!ui.r1c1);
    update('grid', ui.showGridLines !== false);
    update('headers', ui.showHeaders !== false);
  };
  viewBtn?.addEventListener('click', (e) => {
    e.stopPropagation();
    if (!viewDrop) return;
    refreshViewMenu();
    viewDrop.hidden = !viewDrop.hidden;
    viewBtn.setAttribute('aria-expanded', viewDrop.hidden ? 'false' : 'true');
  });
  document.addEventListener('mousedown', (e) => {
    if (!viewDrop || viewDrop.hidden) return;
    if (viewDrop.contains(e.target as Node) || viewBtn?.contains(e.target as Node)) return;
    closeViewMenu();
  });
  viewDrop?.querySelectorAll<HTMLButtonElement>('[data-view]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const inst = getInst();
      if (!inst) return;
      const action = btn.dataset.view;
      const ui = inst.store.getState().ui;
      if (action === 'show-formulas') mutators.setShowFormulas(inst.store, !ui.showFormulas);
      else if (action === 'r1c1') mutators.setR1C1(inst.store, !ui.r1c1);
      else if (action === 'grid') mutators.setShowGridLines(inst.store, !ui.showGridLines);
      else if (action === 'headers') mutators.setShowHeaders(inst.store, !ui.showHeaders);
      refreshViewMenu();
    });
  });

  // ── Tools menu (Iterative / Names / Conditional) ─────────────────────────
  const toolsBtn = document.getElementById('menu-tools');
  const toolsDrop = document.getElementById('menu-tools-dropdown');
  const closeToolsMenu = (): void => {
    if (!toolsDrop) return;
    toolsDrop.hidden = true;
    toolsBtn?.setAttribute('aria-expanded', 'false');
  };
  toolsBtn?.addEventListener('click', (e) => {
    e.stopPropagation();
    if (!toolsDrop) return;
    toolsDrop.hidden = !toolsDrop.hidden;
    toolsBtn.setAttribute('aria-expanded', toolsDrop.hidden ? 'false' : 'true');
  });
  document.addEventListener('mousedown', (e) => {
    if (!toolsDrop || toolsDrop.hidden) return;
    if (toolsDrop.contains(e.target as Node) || toolsBtn?.contains(e.target as Node)) return;
    closeToolsMenu();
  });
  toolsDrop?.querySelectorAll<HTMLButtonElement>('[data-tools]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const inst = getInst();
      if (!inst) return;
      const action = btn.dataset.tools;
      closeToolsMenu();
      if (action === 'iterative') inst.openIterativeDialog();
      else if (action === 'named') inst.openNamedRangeDialog();
      else if (action === 'conditional') inst.openConditionalDialog();
    });
  });

  // ── Merge / Wrap / Sort buttons ───────────────────────────────────────────
  document.getElementById('btn-merge')?.addEventListener('click', () => {
    const inst = getInst();
    if (!inst) return;
    const s = inst.store.getState();
    const r = s.selection.range;
    const anchorAt0 = s.merges.byAnchor.get(`${r.sheet}:${r.r0}:${r.c0}`);
    const isExactMerge =
      anchorAt0 &&
      r.r0 === anchorAt0.r0 &&
      r.c0 === anchorAt0.c0 &&
      r.r1 === anchorAt0.r1 &&
      r.c1 === anchorAt0.c1;
    if (isExactMerge) applyUnmerge(inst.store, inst.workbook, inst.history, r);
    else applyMerge(inst.store, inst.workbook, inst.history, r);
    sheetEl.focus();
  });

  document.getElementById('btn-wrap')?.addEventListener('click', () => {
    const inst = getInst();
    if (!inst) return;
    const current = inst;
    recordFormatChange(inst.history, inst.store, () => {
      toggleWrap(current.store.getState(), current.store);
    });
    sheetEl.focus();
  });

  return { refreshViewMenu };
};
