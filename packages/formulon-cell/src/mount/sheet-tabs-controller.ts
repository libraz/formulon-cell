import type { History } from '../commands/history.js';
import { moveSheet, removeSheet, setSheetHidden } from '../commands/sheet-mutate.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import { inheritHostTokens } from '../interact/inherit-host-tokens.js';
import type { SpreadsheetStore } from '../store/store.js';
import { mutators } from '../store/store.js';
import { hiddenSheetIndexes, visibleSheetIndexes } from './sheet-indexes.js';
import {
  createSheetMenuButton,
  createSheetMenuSeparator,
  formatSheetLabel,
  positionSheetMenu,
} from './sheet-menu.js';

interface SheetTabsControllerInput {
  addSheetBtn: HTMLButtonElement;
  firstSheet: HTMLButtonElement;
  getStrings: () => Strings;
  getWb: () => WorkbookHandle;
  history: History;
  host: HTMLElement;
  hydrateActiveSheet: () => void;
  invalidate: () => void;
  lastSheet: HTMLButtonElement;
  refreshStatusBar: () => void;
  sheetMenu: HTMLDivElement;
  sheetTabs: HTMLDivElement;
  store: SpreadsheetStore;
}

export interface SheetTabsController {
  closeMenu(): void;
  detach(): void;
  showMenu(idx: number, tab: HTMLButtonElement, x: number, y: number): void;
  switchRelative(delta: 1 | -1): void;
  switchSheet(idx: number): void;
  update(): void;
}

export function attachSheetTabsController(input: SheetTabsControllerInput): SheetTabsController {
  const {
    addSheetBtn,
    firstSheet,
    getStrings,
    getWb,
    history,
    host,
    hydrateActiveSheet,
    invalidate,
    lastSheet,
    refreshStatusBar,
    sheetMenu,
    sheetTabs,
    store,
  } = input;

  const visible = (): number[] => visibleSheetIndexes(getWb(), store);
  const hidden = (): number[] => hiddenSheetIndexes(getWb(), store);

  const refreshLabels = (): void => {
    const t = getStrings().sheetTabs;
    firstSheet.setAttribute('aria-label', t.previousSheet);
    lastSheet.setAttribute('aria-label', t.nextSheet);
    sheetTabs.setAttribute('aria-label', t.workbookSheets);
    addSheetBtn.setAttribute('aria-label', t.addSheet);
  };

  const closeMenu = (): void => {
    sheetMenu.hidden = true;
    sheetMenu.replaceChildren();
  };

  const switchSheet = (idx: number): void => {
    const wb = getWb();
    if (idx < 0 || idx >= wb.sheetCount) return;
    if (store.getState().layout.hiddenSheets.has(idx)) return;
    if (idx === store.getState().data.sheetIndex) return;
    wb.clearViewportHint();
    mutators.setSheetIndex(store, idx);
    hydrateActiveSheet();
    update();
    refreshStatusBar();
    invalidate();
  };

  const switchRelative = (delta: 1 | -1): void => {
    const indexes = visible();
    const pos = indexes.indexOf(store.getState().data.sheetIndex);
    const next = indexes[pos + delta];
    if (next !== undefined) switchSheet(next);
  };

  const beginRename = (idx: number, tab: HTMLButtonElement): void => {
    const wb = getWb();
    const strings = getStrings();
    const before = wb.sheetName(idx);
    const renameInput = document.createElement('input');
    renameInput.type = 'text';
    renameInput.className = 'fc-host__sheetbar-rename';
    renameInput.value = before;
    renameInput.spellcheck = false;
    renameInput.autocomplete = 'off';
    renameInput.setAttribute('aria-label', formatSheetLabel(strings.sheetTabs.renameSheet, before));
    tab.replaceWith(renameInput);

    let done = false;
    const finish = (commit: boolean): void => {
      if (done) return;
      done = true;
      const next = renameInput.value.trim();
      if (commit && next && next !== before) {
        const ok = wb.renameSheet(idx, next);
        if (!ok) {
          renameInput.setAttribute('aria-invalid', 'true');
          done = false;
          renameInput.focus();
          renameInput.select();
          return;
        }
      }
      update();
    };

    renameInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        finish(true);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        finish(false);
      }
    });
    renameInput.addEventListener('blur', () => finish(true));
    requestAnimationFrame(() => {
      renameInput.focus();
      renameInput.select();
    });
  };

  const showMenu = (idx: number, tab: HTMLButtonElement, x: number, y: number): void => {
    const wb = getWb();
    const strings = getStrings();
    const visibleIndexes = visible();
    const hiddenIndexes = hidden();
    const canMutate = wb.capabilities.sheetMutate;
    const canHide = wb.capabilities.sheetTabHidden;
    const moveAndRefresh = (from: number, to: number): void => {
      if (!moveSheet(store, wb, from, to)) return;
      hydrateActiveSheet();
      update();
      refreshStatusBar();
      invalidate();
    };
    const menuButton = (label: string, onClick: () => void, disabled = false): HTMLButtonElement =>
      createSheetMenuButton(label, onClick, closeMenu, disabled);

    sheetMenu.replaceChildren(
      menuButton(strings.sheetTabs.rename, () => beginRename(idx, tab), !canMutate),
      menuButton(strings.sheetTabs.insertSheet, () => {
        const added = wb.addSheet();
        if (added < 0) return;
        switchSheet(added);
        update();
      }),
      menuButton(
        strings.sheetTabs.moveLeft,
        () => moveAndRefresh(idx, idx - 1),
        !canMutate || idx <= 0,
      ),
      menuButton(
        strings.sheetTabs.moveRight,
        () => moveAndRefresh(idx, idx + 1),
        !canMutate || idx >= wb.sheetCount - 1,
      ),
      createSheetMenuSeparator(),
      menuButton(
        strings.sheetTabs.deleteSheet,
        () => {
          if (!removeSheet(store, wb, idx)) return;
          hydrateActiveSheet();
          update();
          refreshStatusBar();
          invalidate();
        },
        !canMutate || wb.sheetCount <= 1,
      ),
      menuButton(
        strings.sheetTabs.hideSheet,
        () => {
          if (!setSheetHidden(store, wb, history, idx, true)) return;
          hydrateActiveSheet();
          update();
          refreshStatusBar();
          invalidate();
        },
        !canHide || visibleIndexes.length <= 1,
      ),
      menuButton(
        hiddenIndexes.length > 0
          ? formatSheetLabel(
              strings.sheetTabs.unhideNamedSheet,
              wb.sheetName(hiddenIndexes[0] ?? 0),
            )
          : strings.sheetTabs.unhideSheet,
        () => {
          const target = hiddenIndexes[0];
          if (target === undefined) return;
          if (!setSheetHidden(store, wb, history, target, false)) return;
          switchSheet(target);
          update();
        },
        !canHide || hiddenIndexes.length === 0,
      ),
    );
    inheritHostTokens(host, sheetMenu);
    positionSheetMenu(sheetMenu, x, y);
    sheetMenu.querySelector<HTMLButtonElement>('.fc-sheetmenu__item:not([disabled])')?.focus();
  };

  const update = (): void => {
    const wb = getWb();
    const active = store.getState().data.sheetIndex;
    const visibleIndexes = visible();
    refreshLabels();
    sheetTabs.replaceChildren();
    for (const idx of visibleIndexes) {
      const tab = document.createElement('button');
      tab.type = 'button';
      tab.className = 'fc-host__sheetbar-tab';
      tab.setAttribute('role', 'tab');
      tab.dataset.fcSheetIndex = String(idx);
      const selected = idx === active;
      tab.setAttribute('aria-selected', selected ? 'true' : 'false');
      tab.tabIndex = selected ? 0 : -1;
      tab.textContent = wb.sheetName(idx);
      tab.addEventListener('click', () => switchSheet(idx));
      tab.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        e.stopPropagation();
        switchSheet(idx);
        showMenu(idx, tab, e.clientX, e.clientY);
      });
      tab.addEventListener('dblclick', () => beginRename(idx, tab));
      tab.addEventListener('keydown', (e) => {
        if (e.key === 'F2' || e.key === 'Enter') {
          e.preventDefault();
          beginRename(idx, tab);
        } else if (e.key === 'ContextMenu' || (e.shiftKey && e.key === 'F10')) {
          e.preventDefault();
          const r = tab.getBoundingClientRect();
          showMenu(idx, tab, r.left, r.bottom + 2);
        } else if (e.key === 'ArrowLeft') {
          e.preventDefault();
          switchRelative(-1);
        } else if (e.key === 'ArrowRight') {
          e.preventDefault();
          switchRelative(1);
        }
      });
      sheetTabs.appendChild(tab);
      if (selected) {
        requestAnimationFrame(() => {
          if (tab.isConnected) tab.scrollIntoView({ block: 'nearest', inline: 'nearest' });
        });
      }
    }
    const activePos = visibleIndexes.indexOf(active);
    firstSheet.disabled = activePos <= 0;
    lastSheet.disabled = activePos < 0 || activePos >= visibleIndexes.length - 1;
    addSheetBtn.disabled = false;
  };

  const onFirstSheetClick = (): void => switchRelative(-1);
  const onLastSheetClick = (): void => switchRelative(1);
  const onAddSheetClick = (): void => {
    const idx = getWb().addSheet();
    if (idx < 0) return;
    switchSheet(idx);
    update();
  };
  const onSheetMenuPointerDown = (e: PointerEvent): void => {
    if (sheetMenu.hidden) return;
    const target = e.target;
    if (target instanceof Node && sheetMenu.contains(target)) return;
    closeMenu();
  };
  const onSheetMenuKeyDown = (e: KeyboardEvent): void => {
    if (sheetMenu.hidden) return;
    if (e.key === 'Escape') closeMenu();
  };

  firstSheet.addEventListener('click', onFirstSheetClick);
  lastSheet.addEventListener('click', onLastSheetClick);
  addSheetBtn.addEventListener('click', onAddSheetClick);
  document.addEventListener('pointerdown', onSheetMenuPointerDown);
  document.addEventListener('keydown', onSheetMenuKeyDown);

  return {
    closeMenu,
    detach(): void {
      firstSheet.removeEventListener('click', onFirstSheetClick);
      lastSheet.removeEventListener('click', onLastSheetClick);
      addSheetBtn.removeEventListener('click', onAddSheetClick);
      document.removeEventListener('pointerdown', onSheetMenuPointerDown);
      document.removeEventListener('keydown', onSheetMenuKeyDown);
      sheetMenu.remove();
    },
    showMenu,
    switchRelative,
    switchSheet,
    update,
  };
}
