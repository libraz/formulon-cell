import { canExecuteBuiltIn } from '../commands/built-in-command-policy.js';
import type { History } from '../commands/history.js';
import type { InteractionOrigin } from '../commands/interaction-policy.js';
import { isSheetProtected, isWorkbookStructureProtected } from '../commands/protection.js';
import {
  copySheetAction,
  deleteSheetAction,
  insertSheetAction,
  moveSheetAction,
  renameSheetAction,
  type SheetActionContext,
  setSheetHiddenAction,
  setSheetTabColorAction,
  setSheetVisibilityAction,
} from '../commands/sheet-actions.js';
import {
  createColorPalette,
  pickNativeColor,
  themeColorColumns,
} from '../components/color-palette.js';
import { isMacWorkbook } from '../engine/spreadsheet-profile.js';
import { SheetVisibility } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import { clampPanelToViewport } from '../interact/overlay-position.js';
import { SHEET_TAB_COLOR_CHOICES, sheetTabColorChoiceLabel } from '../sheet-tab-colors.js';
import type { SpreadsheetStore } from '../store/store.js';
import { mutators } from '../store/store.js';
import { showMoveCopySheetDialog } from '../toolbar/dialogs/move-copy-sheet.js';
import { showMessage } from '../toolbar/dialogs/prompt.js';
import { confirmSheetDelete } from '../toolbar/dialogs/sheet-delete-confirm.js';
import { showUnhideSheetDialog } from '../toolbar/dialogs/unhide-sheet.js';
import { onPointerArrive, projectDisabledState } from '../toolbar/menu-a11y.js';
import { hiddenSheetIndexes, visibleSheetIndexes } from './sheet-indexes.js';
import {
  createSheetMenuButton,
  createSheetMenuColorButton,
  createSheetMenuSeparator,
  createSheetMenuSubmenuTrigger,
  createSheetTabButton,
  formatSheetLabel,
  positionSheetMenu,
} from './sheet-menu.js';
import { attachSheetTabDrag } from './sheet-tab-drag.js';

const SHEET_NAV_REPEAT_DELAY_MS = 350;
const SHEET_NAV_REPEAT_INTERVAL_MS = 120;

interface SheetTabsControllerInput {
  addSheetBtn: HTMLButtonElement;
  firstSheet: HTMLButtonElement;
  getStrings: () => Strings;
  /** Name shown as the target book in Move or Copy; blank when the host gives none. */
  getWorkbookName?: () => string | null | undefined;
  getWb: () => WorkbookHandle;
  history: History;
  hydrateActiveSheet: () => void;
  invalidate: () => void;
  lastSheet: HTMLButtonElement;
  refreshStatusBar: () => void;
  sheetMenu: HTMLDivElement;
  sheetTabs: HTMLDivElement;
  store: SpreadsheetStore;
  /** Opens the protect or unprotect flow of the active sheet. */
  toggleSheetProtection?: () => void;
}

export interface SheetTabsController {
  /** Insert a sheet from `origin` (the tab bar's own button when absent) and switch to it. */
  addSheet(origin?: InteractionOrigin): void;
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
  // The tab bar, its menu and its gestures are one entry point; the keyboard passes its own.
  const sheetCtx = (origin: InteractionOrigin = 'contextMenu'): SheetActionContext => ({
    store,
    workbook: getWb(),
    history,
    origin,
  });
  const denied = (id: Parameters<typeof canExecuteBuiltIn>[1]): boolean =>
    !canExecuteBuiltIn(store, id, 'contextMenu').allowed;
  let sheetMenuRestoreFocus: HTMLElement | null = null;

  // The Mac arrows scroll the tab strip by one tab; elsewhere they switch sheets.
  const scrollsStrip = (): boolean => isMacWorkbook(getWb());
  const tabLeft = (tab: HTMLElement): number =>
    tab.getBoundingClientRect().left -
    sheetTabs.getBoundingClientRect().left +
    sheetTabs.scrollLeft;
  const stripAtStart = (): boolean => sheetTabs.scrollLeft <= 0;
  const stripAtEnd = (): boolean =>
    sheetTabs.scrollLeft + sheetTabs.clientWidth >= sheetTabs.scrollWidth - 1;
  const scrollStrip = (delta: 1 | -1): void => {
    const lefts = Array.from(sheetTabs.querySelectorAll<HTMLElement>('.fc-host__sheetbar-tab')).map(
      tabLeft,
    );
    const here = sheetTabs.scrollLeft;
    const next =
      delta > 0 ? lefts.find((l) => l > here + 1) : [...lefts].reverse().find((l) => l < here - 1);
    sheetTabs.scrollLeft = delta > 0 ? (next ?? sheetTabs.scrollWidth) : (next ?? 0);
    refreshNav();
  };
  const step = (delta: 1 | -1): void => {
    if (scrollsStrip()) scrollStrip(delta);
    else switchRelative(delta);
  };

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
      step(delta);
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
      step(delta);
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

  const refreshAfterSheetChange = (): void => {
    hydrateActiveSheet();
    update();
    refreshStatusBar();
    invalidate();
  };

  const switchSheet = (idx: number): void => {
    const wb = getWb();
    if (idx < 0 || idx >= wb.sheetCount) return;
    if (store.getState().layout.hiddenSheets.has(idx)) return;
    if (idx === store.getState().data.sheetIndex) return;
    mutators.setSheetIndex(store, idx);
    refreshAfterSheetChange();
  };

  const switchRelative = (delta: 1 | -1): void => {
    const indexes = visible();
    const pos = indexes.indexOf(store.getState().data.sheetIndex);
    const next = indexes[pos + delta];
    if (next !== undefined) switchSheet(next);
  };

  const beginRename = (idx: number, tab: HTMLButtonElement): void => {
    if (denied('renameSheet')) return;
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
    // The editor stands in for the tab at the tab's own width.
    const tabWidth = tab.getBoundingClientRect().width;
    if (tabWidth > 0) renameInput.style.width = `${tabWidth}px`;
    tab.replaceWith(renameInput);

    let done = false;
    const finish = (commit: boolean): void => {
      if (done) return;
      done = true;
      const next = renameInput.value.trim();
      if (commit && next && next !== before) {
        const ok = renameSheetAction(sheetCtx(), idx, next);
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
    // Very-hidden is the only tab state `setSheetTabHidden` cannot express, so
    // offering it without the three-state setter would promise a state the
    // saved file could not carry.
    const canVeryHide = canHide && wb.capabilities.sheetVisibility === true;
    const moveAndRefresh = (from: number, to: number): void => {
      if (!moveSheetAction(sheetCtx(), from, to)) return;
      refreshAfterSheetChange();
    };
    const openMoveCopy = async (): Promise<void> => {
      const shown = visible();
      const result = await showMoveCopySheetDialog({
        host: sheetTabs,
        strings: {
          title: strings.sheetTabs.moveCopyTitle,
          targetBook: strings.sheetTabs.moveCopyTargetBook,
          before: strings.sheetTabs.moveCopyBefore,
          toEnd: strings.sheetTabs.moveCopyToEnd,
          makeCopy: strings.sheetTabs.moveCopyMakeCopy,
          ok: strings.hyperlinkDialog.ok,
          cancel: strings.hyperlinkDialog.cancel,
        },
        bookName: input.getWorkbookName?.() ?? '',
        sheets: shown.map((i) => wb.sheetName(i)),
        current: Math.max(0, shown.indexOf(idx)),
        copyDisabledReason: denied('copySheet') || !canMutate ? sheetActionReason() : null,
      });
      if (!result) return;
      const beforeSheet = shown[result.before];
      if (result.copy) {
        const copy = copySheetAction(sheetCtx(), idx, beforeSheet ?? wb.sheetCount);
        if (copy < 0) {
          await showMessage({
            host: sheetTabs,
            title: strings.sheetTabs.moveCopyTitle,
            message: strings.sheetTabs.moveCopyFailed,
            okLabel: strings.hyperlinkDialog.ok,
          });
          return;
        }
        refreshAfterSheetChange();
        return;
      }
      const last = shown.at(-1) ?? idx;
      const to =
        beforeSheet === undefined ? last : idx < beforeSheet ? beforeSheet - 1 : beforeSheet;
      moveAndRefresh(idx, Math.max(0, Math.min(wb.sheetCount - 1, to)));
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
    const veryHideReason = (): string =>
      structureProtected || !canVeryHide
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
            setSheetTabColorAction(sheetCtx(), idx, choice.color);
            update();
            closeMenu();
          },
        ),
      ),
    );
    for (const swatch of swatches.querySelectorAll<HTMLButtonElement>('button')) {
      projectDisabledState(swatch, denied('sheetTabColor'), sheetActionReason());
    }
    colorPalette.append(colorLabel, swatches);
    const unhideButtons =
      hiddenIndexes.length > 0
        ? hiddenIndexes.map((target) =>
            menuButton(
              formatSheetLabel(strings.sheetTabs.unhideNamedSheet, wb.sheetName(target)),
              () => {
                if (!setSheetHiddenAction(sheetCtx(), target, false)) return;
                switchSheet(target);
                update();
              },
              denied('unhideSheet') || !canHide,
              unhideReason(true),
            ),
          )
        : [menuButton(strings.sheetTabs.unhideSheet, () => undefined, true, unhideReason(false))];

    const insertSheetFromMenu = (): void => {
      const added = insertSheetAction(sheetCtx());
      if (added < 0) return;
      switchSheet(added);
      update();
    };
    const deleteSheetFromMenu = async (): Promise<void> => {
      if (denied('deleteSheet') || !(await confirmSheetDelete(strings, wb, idx))) return;
      if (!deleteSheetAction(sheetCtx(), idx)) return;
      refreshAfterSheetChange();
    };
    const hideSheetAction = (): void => {
      if (!setSheetHiddenAction(sheetCtx(), idx, true)) return;
      refreshAfterSheetChange();
    };
    const desktopItems: HTMLElement[] = [
      menuButton(
        strings.sheetTabs.rename,
        () => beginRename(idx, tab),
        denied('renameSheet') || !canMutate,
        sheetActionReason(),
      ),
      menuButton(
        strings.sheetTabs.insertSheet,
        insertSheetFromMenu,
        denied('insertSheet'),
        strings.ribbonMenu.workbookStructureProtectedBlocked,
      ),
      menuButton(
        strings.sheetTabs.moveLeft,
        () => moveAndRefresh(idx, idx - 1),
        denied('moveSheet') || !canMutate || idx <= 0,
        moveReason(idx <= 0),
      ),
      menuButton(
        strings.sheetTabs.moveRight,
        () => moveAndRefresh(idx, idx + 1),
        denied('moveSheet') || !canMutate || idx >= wb.sheetCount - 1,
        moveReason(idx >= wb.sheetCount - 1),
      ),
      menuButton(
        strings.sheetTabs.moveCopy,
        () => {
          void openMoveCopy();
        },
        denied('moveSheet') || !canMutate,
        sheetActionReason(),
      ),
      createSheetMenuSeparator(),
      colorPalette,
      createSheetMenuSeparator(),
      menuButton(
        strings.sheetTabs.deleteSheet,
        deleteSheetFromMenu,
        denied('deleteSheet') || !canMutate || wb.sheetCount <= 1,
        deleteReason(),
      ),
      menuButton(
        strings.sheetTabs.hideSheet,
        hideSheetAction,
        denied('hideSheet') || !canHide || visibleIndexes.length <= 1,
        hideReason(),
      ),
      menuButton(
        strings.sheetTabs.veryHideSheet,
        () => {
          if (!setSheetVisibilityAction(sheetCtx(), idx, SheetVisibility.VeryHidden)) {
            return;
          }
          refreshAfterSheetChange();
        },
        denied('hideSheet') || !canVeryHide || visibleIndexes.length <= 1,
        veryHideReason(),
      ),
      ...unhideButtons,
    ];
    // The Mac tab-colour submenu opens beside the menu with the shared palette.
    const openColorSub = (trigger: HTMLButtonElement): void => {
      const m = strings.mac.sheetTabMenu;
      const sub = document.createElement('div');
      sub.className = 'fc-sheetmenu__sub';
      sub.setAttribute('role', 'menu');
      sub.setAttribute('aria-label', m.tabColor);
      const apply = (color: string | null): void => {
        setSheetTabColorAction(sheetCtx(), idx, color);
        update();
        closeMenu();
      };
      const palette = createColorPalette({
        themeLabel: strings.ribbon.themeColors,
        standardLabel: strings.ribbon.standardColors,
        moreColorsLabel: strings.ribbon.moreColors,
        highContrastOnlyLabel: strings.ribbon.highContrastOnly,
        noFillLabel: strings.sheetTabs.noColor,
        ariaLabel: m.tabColor,
        theme: themeColorColumns(wb.getTheme?.()?.colors),
        value: tabColor ?? null,
        onPick: (color) => apply(color || null),
        onMoreColors: () => pickNativeColor(tabColor ?? '#ffffff', apply),
      });
      sub.appendChild(palette.el);
      sheetMenu.appendChild(sub);
      const menuRect = sheetMenu.getBoundingClientRect();
      const rowRect = trigger.getBoundingClientRect();
      const { x: left, y: top } = clampPanelToViewport(sub, menuRect.right - 1, rowRect.top, {
        pad: 8,
        fallbackWidth: 198,
        fallbackHeight: 285,
      });
      sub.style.left = `${left}px`;
      sub.style.top = `${top}px`;
      trigger.setAttribute('aria-expanded', 'true');
    };
    // The Mac desktop's tab menu lists its own entries in its own order.
    const macItems = (): HTMLElement[] => {
      const m = strings.mac.sheetTabMenu;
      const protectedSheet = isSheetProtected(store.getState(), idx);
      const sheetProtection = input.toggleSheetProtection;
      const colorTrigger = createSheetMenuSubmenuTrigger(m.tabColor);
      const toggleColorSub = (open: boolean): void => {
        sheetMenu.querySelector('.fc-sheetmenu__sub')?.remove();
        colorTrigger.setAttribute('aria-expanded', 'false');
        if (open) openColorSub(colorTrigger);
      };
      colorTrigger.addEventListener('click', () => toggleColorSub(true));
      onPointerArrive(colorTrigger, () => toggleColorSub(true));
      colorTrigger.addEventListener('keydown', (e) => {
        if (e.key !== 'ArrowRight') return;
        e.preventDefault();
        toggleColorSub(true);
        sheetMenu.querySelector<HTMLButtonElement>('.fc-sheetmenu__sub button')?.focus();
      });
      const unhideSheetsAction = async (): Promise<void> => {
        const targets = hidden();
        const pos = await showUnhideSheetDialog({
          host: sheetTabs,
          strings: {
            title: m.unhideTitle,
            list: m.unhideList,
            ok: strings.hyperlinkDialog.ok,
            cancel: strings.hyperlinkDialog.cancel,
          },
          sheets: targets.map((i) => wb.sheetName(i)),
        });
        const target = pos === null ? undefined : targets[pos];
        if (target === undefined || !setSheetHiddenAction(sheetCtx(), target, false)) return;
        switchSheet(target);
        update();
      };
      return [
        menuButton(
          m.insertSheet,
          insertSheetFromMenu,
          denied('insertSheet'),
          strings.ribbonMenu.workbookStructureProtectedBlocked,
        ),
        menuButton(
          strings.sheetTabs.deleteSheet,
          deleteSheetFromMenu,
          denied('deleteSheet') || !canMutate || wb.sheetCount <= 1,
          deleteReason(),
        ),
        menuButton(
          strings.sheetTabs.rename,
          () => beginRename(idx, tab),
          denied('renameSheet') || !canMutate,
          sheetActionReason(),
        ),
        menuButton(
          strings.sheetTabs.moveCopy,
          () => {
            void openMoveCopy();
          },
          denied('moveSheet') || !canMutate,
          sheetActionReason(),
        ),
        menuButton(m.viewCode, () => undefined, true, m.unavailable),
        menuButton(
          protectedSheet
            ? strings.ribbonMenu.unprotectSheetCommand
            : strings.ribbonMenu.protectSheetCommand,
          () => sheetProtection?.(),
          !sheetProtection || denied('protect'),
          m.unavailable,
        ),
        colorTrigger,
        createSheetMenuSeparator(),
        menuButton(
          m.hideSheet,
          hideSheetAction,
          denied('hideSheet') || !canHide || visibleIndexes.length <= 1,
          hideReason(),
        ),
        menuButton(
          m.unhideSheet,
          () => {
            void unhideSheetsAction();
          },
          denied('unhideSheet') || !canHide || hiddenIndexes.length === 0,
          unhideReason(hiddenIndexes.length > 0),
        ),
        createSheetMenuSeparator(),
        menuButton(m.selectAllSheets, () => undefined, true, m.unavailable),
      ];
    };
    sheetMenu.replaceChildren(...(isMacWorkbook(wb) ? macItems() : desktopItems));
    for (const entry of sheetMenu.querySelectorAll<HTMLElement>('.fc-sheetmenu__item')) {
      if (entry.getAttribute('aria-haspopup')) continue;
      onPointerArrive(entry, () => {
        sheetMenu.querySelector('.fc-sheetmenu__sub')?.remove();
        sheetMenu.querySelector('[aria-haspopup="menu"]')?.setAttribute('aria-expanded', 'false');
      });
    }
    positionSheetMenu(sheetMenu, x, y);
    sheetMenu.querySelector<HTMLButtonElement>('.fc-sheetmenu__item:not([disabled])')?.focus();
  };

  const refreshNav = (): void => {
    const strings = getStrings();
    const indexes = visible();
    const activePos = indexes.indexOf(store.getState().data.sheetIndex);
    const strip = scrollsStrip();
    setSheetButtonDisabled(
      firstSheet,
      strip ? stripAtStart() : activePos <= 0,
      strings.sheetTabs.previousSheetUnavailable,
      strings.sheetTabs.previousSheet,
    );
    setSheetButtonDisabled(
      lastSheet,
      strip ? stripAtEnd() : activePos < 0 || activePos >= indexes.length - 1,
      strings.sheetTabs.nextSheetUnavailable,
      strings.sheetTabs.nextSheet,
    );
  };

  const update = (): void => {
    const wb = getWb();
    const active = store.getState().data.sheetIndex;
    const visibleIndexes = visible();
    refreshLabels();
    sheetTabs.dataset.fcSheetOverflow = visibleIndexes.length > 1 ? 'true' : 'false';
    sheetTabs.replaceChildren();
    // The Mac desktop marks a protected sheet's tab with a padlock.
    const mac = isMacWorkbook(wb);
    for (const idx of visibleIndexes) {
      const selected = idx === active;
      const tabColor = store.getState().layout.sheetTabColors.get(idx);
      const tab = createSheetTabButton({
        index: idx,
        label: wb.sheetName(idx),
        selected,
        tabColor,
        locked: mac && isSheetProtected(store.getState(), idx),
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
    refreshNav();
    setSheetButtonDisabled(
      addSheetBtn,
      denied('insertSheet'),
      strings.ribbonMenu.workbookStructureProtectedBlocked,
      strings.sheetTabs.addSheet,
    );
  };

  const firstSheetRepeater = createSheetNavRepeater(firstSheet, -1);
  const lastSheetRepeater = createSheetNavRepeater(lastSheet, 1);
  const addAndSwitchSheet = (origin?: InteractionOrigin): void => {
    // The Mac desktop inserts the new sheet after the active one.
    const idx = insertSheetAction(sheetCtx(origin), isMacWorkbook(getWb()));
    if (idx < 0) return;
    switchSheet(idx);
    update();
  };
  const onAddSheetClick = (): void => addAndSwitchSheet();
  const onSheetMenuPointerDown = (e: PointerEvent): void => {
    if (sheetMenu.hidden) return;
    const target = e.target;
    if (target instanceof Node && sheetMenu.contains(target)) return;
    closeMenu();
  };
  const onSheetMenuFocusIn = (e: FocusEvent): void => {
    if (sheetMenu.hidden) return;
    const target = e.target;
    if (target instanceof Node && sheetMenu.contains(target)) return;
    closeMenu();
  };
  const onSheetMenuKeyDown = (e: KeyboardEvent): void => {
    if (sheetMenu.hidden) return;
    if (!sheetMenu.contains(document.activeElement)) {
      closeMenu();
      return;
    }
    if (e.key === 'Tab') {
      closeMenu();
      return;
    }
    const inSub = document.activeElement?.closest('.fc-sheetmenu__sub');
    if (inSub) {
      // The palette owns its arrows; Left and Escape fold the submenu back.
      if (e.key === 'Escape' || e.key === 'ArrowLeft') {
        e.preventDefault();
        e.stopPropagation();
        sheetMenu.querySelector('.fc-sheetmenu__sub')?.remove();
        const trigger = sheetMenu.querySelector<HTMLButtonElement>('[aria-haspopup="menu"]');
        trigger?.setAttribute('aria-expanded', 'false');
        trigger?.focus();
      }
      return;
    }
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
      e.stopPropagation();
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
  sheetTabs.addEventListener('scroll', refreshNav);
  const tabDrag = attachSheetTabDrag({
    sheetTabs,
    onDrop: (from, to) => {
      if (!getWb().capabilities.sheetMutate) return;
      if (!moveSheetAction(sheetCtx(), from, to)) return;
      // The dragged tab becomes the active sheet.
      if (store.getState().data.sheetIndex !== to) mutators.setSheetIndex(store, to);
      refreshAfterSheetChange();
    },
  });
  document.addEventListener('pointerdown', onSheetMenuPointerDown);
  document.addEventListener('focusin', onSheetMenuFocusIn, true);
  document.addEventListener('keydown', onSheetMenuKeyDown, true);

  return {
    addSheet: addAndSwitchSheet,
    closeMenu,
    detach(): void {
      firstSheetRepeater.detach();
      lastSheetRepeater.detach();
      firstSheet.removeEventListener('click', firstSheetRepeater.onClick);
      lastSheet.removeEventListener('click', lastSheetRepeater.onClick);
      addSheetBtn.removeEventListener('click', onAddSheetClick);
      sheetTabs.removeEventListener('scroll', refreshNav);
      tabDrag.detach();
      document.removeEventListener('pointerdown', onSheetMenuPointerDown);
      document.removeEventListener('focusin', onSheetMenuFocusIn, true);
      document.removeEventListener('keydown', onSheetMenuKeyDown, true);
      sheetMenu.remove();
    },
    showMenu,
    switchRelative,
    switchSheet,
    update,
  };
}
