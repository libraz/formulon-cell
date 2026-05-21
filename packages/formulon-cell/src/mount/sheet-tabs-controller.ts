import type { History } from '../commands/history.js';
import { isWorkbookStructureProtected } from '../commands/protection.js';
import {
  addSheet,
  moveSheet,
  removeSheet,
  renameSheet,
  setSheetHidden,
} from '../commands/sheet-mutate.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import { inheritHostTokens } from '../interact/inherit-host-tokens.js';
import { SHEET_TAB_COLOR_CHOICES, sheetTabColorChoiceLabel } from '../sheet-tab-colors.js';
import type { SpreadsheetStore } from '../store/store.js';
import { mutators } from '../store/store.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { hiddenSheetIndexes, visibleSheetIndexes } from './sheet-indexes.js';
import {
  createSheetMenuButton,
  createSheetMenuColorButton,
  createSheetMenuSeparator,
  createSheetTabButton,
  formatSheetLabel,
  positionSheetMenu,
} from './sheet-menu.js';

const SHEET_NAV_REPEAT_DELAY_MS = 350;
const SHEET_NAV_REPEAT_INTERVAL_MS = 120;

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
  let sheetMenuRestoreFocus: HTMLElement | null = null;

  const createSheetNavRepeater = (
    button: HTMLButtonElement,
    delta: 1 | -1,
  ): {
    detach(): void;
    onClick(e: MouseEvent): void;
  } => {
    let delayTimer: ReturnType<typeof setTimeout> | undefined;
    let intervalTimer: ReturnType<typeof setInterval> | undefined;
    let suppressNextClick = false;

    const clearTimers = (): void => {
      if (delayTimer !== undefined) {
        clearTimeout(delayTimer);
        delayTimer = undefined;
      }
      if (intervalTimer !== undefined) {
        clearInterval(intervalTimer);
        intervalTimer = undefined;
      }
    };
    const repeat = (): void => {
      if (button.disabled || !button.isConnected) {
        clearTimers();
        return;
      }
      suppressNextClick = true;
      switchRelative(delta);
    };
    const onPointerDown = (e: PointerEvent): void => {
      if (e.button !== 0 || button.disabled) return;
      clearTimers();
      delayTimer = setTimeout(() => {
        delayTimer = undefined;
        repeat();
        intervalTimer = setInterval(repeat, SHEET_NAV_REPEAT_INTERVAL_MS);
      }, SHEET_NAV_REPEAT_DELAY_MS);
    };
    const onPointerUpOrCancel = (): void => {
      clearTimers();
    };
    const onClick = (e: MouseEvent): void => {
      if (suppressNextClick) {
        suppressNextClick = false;
        e.preventDefault();
        e.stopPropagation();
        return;
      }
      switchRelative(delta);
    };

    button.addEventListener('pointerdown', onPointerDown);
    button.addEventListener('pointerup', onPointerUpOrCancel);
    button.addEventListener('pointercancel', onPointerUpOrCancel);
    button.addEventListener('pointerleave', onPointerUpOrCancel);
    button.addEventListener('lostpointercapture', onPointerUpOrCancel);
    return {
      detach(): void {
        clearTimers();
        button.removeEventListener('pointerdown', onPointerDown);
        button.removeEventListener('pointerup', onPointerUpOrCancel);
        button.removeEventListener('pointercancel', onPointerUpOrCancel);
        button.removeEventListener('pointerleave', onPointerUpOrCancel);
        button.removeEventListener('lostpointercapture', onPointerUpOrCancel);
      },
      onClick,
    };
  };

  const refreshLabels = (): void => {
    const t = getStrings().sheetTabs;
    firstSheet.setAttribute('aria-label', t.previousSheet);
    lastSheet.setAttribute('aria-label', t.nextSheet);
    sheetTabs.setAttribute('aria-label', t.workbookSheets);
    addSheetBtn.setAttribute('aria-label', t.addSheet);
  };

  const setSheetButtonDisabled = (
    button: HTMLButtonElement,
    disabled: boolean,
    reason: string | null,
    titlePrefix: string,
  ): void => {
    projectDisabledState(button, disabled, reason, {
      datasetKey: 'disabledReason',
      titlePrefix,
    });
  };

  const closeMenu = (restoreFocus = false): void => {
    const target = restoreFocus ? sheetMenuRestoreFocus : null;
    sheetMenu.hidden = true;
    sheetMenu.replaceChildren();
    sheetMenuRestoreFocus = null;
    if (target?.isConnected) target.focus();
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
    if (isWorkbookStructureProtected(store.getState())) return;
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
        const ok = renameSheet(wb, idx, next, store, history);
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
    sheetMenuRestoreFocus = tab;
    const wb = getWb();
    const strings = getStrings();
    const visibleIndexes = visible();
    const hiddenIndexes = hidden();
    const structureProtected = isWorkbookStructureProtected(store.getState());
    const canMutate = wb.capabilities.sheetMutate;
    const canHide = wb.capabilities.sheetTabHidden;
    const moveAndRefresh = (from: number, to: number): void => {
      if (!moveSheet(store, wb, from, to, history)) return;
      hydrateActiveSheet();
      update();
      refreshStatusBar();
      invalidate();
    };
    const sheetActionReason = (): string =>
      structureProtected
        ? strings.ribbonMenu.workbookStructureProtectedBlocked
        : strings.ribbonMenu.sheetActionUnavailable;
    const moveReason = (atBoundary: boolean): string =>
      structureProtected || !canMutate
        ? sheetActionReason()
        : atBoundary
          ? strings.ribbonMenu.sheetMoveAtBoundary
          : '';
    const deleteReason = (): string =>
      structureProtected
        ? strings.ribbonMenu.workbookStructureProtectedBlocked
        : !canMutate
          ? strings.ribbonMenu.sheetMutationUnavailable
          : strings.ribbonMenu.sheetDeleteRequiresAnotherSheet;
    const hideReason = (): string =>
      structureProtected || !canHide
        ? sheetActionReason()
        : strings.ribbonMenu.sheetHideRequiresVisibleSheet;
    const unhideReason = (hasHiddenTarget: boolean): string =>
      structureProtected || !canHide
        ? sheetActionReason()
        : hasHiddenTarget
          ? ''
          : strings.ribbonMenu.sheetUnhideRequiresHiddenSheet;
    const menuButton = (
      label: string,
      onClick: () => void,
      disabled = false,
      disabledReason: string | null = null,
    ): HTMLButtonElement =>
      createSheetMenuButton(label, onClick, closeMenu, disabled, disabledReason);
    const tabColor = store.getState().layout.sheetTabColors.get(idx);
    const colorPalette = document.createElement('div');
    colorPalette.className = 'fc-sheetmenu__colors';
    colorPalette.setAttribute('role', 'group');
    colorPalette.setAttribute('aria-label', strings.sheetTabs.tabColor);
    const colorLabel = document.createElement('div');
    colorLabel.className = 'fc-sheetmenu__colors-label';
    colorLabel.textContent = strings.sheetTabs.tabColor;
    const swatches = document.createElement('div');
    swatches.className = 'fc-sheetmenu__swatches';
    swatches.append(
      ...SHEET_TAB_COLOR_CHOICES.map((choice) =>
        createSheetMenuColorButton(
          choice.color
            ? `${strings.sheetTabs.tabColor}: ${sheetTabColorChoiceLabel(choice, strings.sheetTabs)}`
            : strings.sheetTabs.noColor,
          choice.color,
          choice.color === null ? tabColor === undefined : tabColor?.toLowerCase() === choice.color,
          () => {
            mutators.setSheetTabColor(store, idx, choice.color);
            update();
            closeMenu();
          },
        ),
      ),
    );
    colorPalette.append(colorLabel, swatches);
    const unhideButtons =
      hiddenIndexes.length > 0
        ? hiddenIndexes.map((target) =>
            menuButton(
              formatSheetLabel(strings.sheetTabs.unhideNamedSheet, wb.sheetName(target)),
              () => {
                if (!setSheetHidden(store, wb, history, target, false)) return;
                switchSheet(target);
                update();
              },
              structureProtected || !canHide,
              unhideReason(true),
            ),
          )
        : [menuButton(strings.sheetTabs.unhideSheet, () => undefined, true, unhideReason(false))];

    sheetMenu.replaceChildren(
      menuButton(
        strings.sheetTabs.rename,
        () => beginRename(idx, tab),
        structureProtected || !canMutate,
        sheetActionReason(),
      ),
      menuButton(
        strings.sheetTabs.insertSheet,
        () => {
          const added = addSheet(store, wb, history);
          if (added < 0) return;
          switchSheet(added);
          update();
        },
        structureProtected,
        strings.ribbonMenu.workbookStructureProtectedBlocked,
      ),
      menuButton(
        strings.sheetTabs.moveLeft,
        () => moveAndRefresh(idx, idx - 1),
        structureProtected || !canMutate || idx <= 0,
        moveReason(idx <= 0),
      ),
      menuButton(
        strings.sheetTabs.moveRight,
        () => moveAndRefresh(idx, idx + 1),
        structureProtected || !canMutate || idx >= wb.sheetCount - 1,
        moveReason(idx >= wb.sheetCount - 1),
      ),
      createSheetMenuSeparator(),
      colorPalette,
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
        structureProtected || !canMutate || wb.sheetCount <= 1,
        deleteReason(),
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
        structureProtected || !canHide || visibleIndexes.length <= 1,
        hideReason(),
      ),
      ...unhideButtons,
    );
    inheritHostTokens(host, sheetMenu);
    positionSheetMenu(sheetMenu, x, y);
    sheetMenu.querySelector<HTMLButtonElement>('.fc-sheetmenu__item:not([disabled])')?.focus();
  };

  const update = (): void => {
    const wb = getWb();
    const active = store.getState().data.sheetIndex;
    const visibleIndexes = visible();
    const activePos = visibleIndexes.indexOf(active);
    refreshLabels();
    sheetTabs.dataset.fcSheetOverflow = visibleIndexes.length > 1 ? 'true' : 'false';
    sheetTabs.replaceChildren();
    for (const idx of visibleIndexes) {
      const selected = idx === active;
      const tabColor = store.getState().layout.sheetTabColors.get(idx);
      const tab = createSheetTabButton({
        index: idx,
        label: wb.sheetName(idx),
        selected,
        tabColor,
      });
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
        } else if (e.key === 'Home') {
          e.preventDefault();
          const first = visible()[0];
          if (first !== undefined) switchSheet(first);
        } else if (e.key === 'End') {
          e.preventDefault();
          const last = visible().at(-1);
          if (last !== undefined) switchSheet(last);
        }
      });
      sheetTabs.appendChild(tab);
      if (selected) {
        requestAnimationFrame(() => {
          if (tab.isConnected) tab.scrollIntoView({ block: 'nearest', inline: 'nearest' });
        });
      }
    }
    const strings = getStrings();
    const structureProtected = isWorkbookStructureProtected(store.getState());
    setSheetButtonDisabled(
      firstSheet,
      activePos <= 0,
      strings.sheetTabs.previousSheetUnavailable,
      strings.sheetTabs.previousSheet,
    );
    setSheetButtonDisabled(
      lastSheet,
      activePos < 0 || activePos >= visibleIndexes.length - 1,
      strings.sheetTabs.nextSheetUnavailable,
      strings.sheetTabs.nextSheet,
    );
    setSheetButtonDisabled(
      addSheetBtn,
      structureProtected,
      strings.ribbonMenu.workbookStructureProtectedBlocked,
      strings.sheetTabs.addSheet,
    );
  };

  const firstSheetRepeater = createSheetNavRepeater(firstSheet, -1);
  const lastSheetRepeater = createSheetNavRepeater(lastSheet, 1);
  const onAddSheetClick = (): void => {
    const idx = addSheet(store, getWb(), history);
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
    const items = Array.from(sheetMenu.querySelectorAll<HTMLButtonElement>('[role^="menuitem"]'));
    const enabled = items.filter((item) => !item.disabled);
    const active =
      document.activeElement instanceof HTMLButtonElement ? document.activeElement : null;
    const activeIndex = active ? enabled.indexOf(active) : -1;
    const focusAt = (idx: number): void => {
      if (enabled.length === 0) return;
      e.preventDefault();
      e.stopPropagation();
      enabled[(idx + enabled.length) % enabled.length]?.focus();
    };
    if (e.key === 'Escape') {
      e.preventDefault();
      closeMenu(true);
    } else if (e.key === 'ArrowDown') {
      focusAt(activeIndex < 0 ? 0 : activeIndex + 1);
    } else if (e.key === 'ArrowUp') {
      focusAt(activeIndex < 0 ? enabled.length - 1 : activeIndex - 1);
    } else if (e.key === 'Home') {
      focusAt(0);
    } else if (e.key === 'End') {
      focusAt(enabled.length - 1);
    } else if (e.key === 'Enter' || e.key === ' ') {
      if (active && sheetMenu.contains(active) && !active.disabled) {
        e.preventDefault();
        e.stopPropagation();
        active.click();
      }
    }
  };

  firstSheet.addEventListener('click', firstSheetRepeater.onClick);
  lastSheet.addEventListener('click', lastSheetRepeater.onClick);
  addSheetBtn.addEventListener('click', onAddSheetClick);
  document.addEventListener('pointerdown', onSheetMenuPointerDown);
  document.addEventListener('keydown', onSheetMenuKeyDown);

  return {
    closeMenu,
    detach(): void {
      firstSheetRepeater.detach();
      lastSheetRepeater.detach();
      firstSheet.removeEventListener('click', firstSheetRepeater.onClick);
      lastSheet.removeEventListener('click', lastSheetRepeater.onClick);
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
