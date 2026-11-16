import {
  clearSelectedContents,
  collectSelectedContentAddresses,
} from '../commands/clear-contents.js';
import type { ClipboardSnapshot } from '../commands/clipboard/snapshot.js';
import { parseTSV } from '../commands/clipboard/tsv.js';
import { clearComment, commentAt, recordCommentChange } from '../commands/comment.js';
import { commentWrite } from '../commands/comment-actions.js';
import { withSelectionFormatOrigin } from '../commands/format.js';
import { tableForCell } from '../commands/format-as-table.js';
import type { History } from '../commands/history.js';
import { hyperlinkAt, isFollowableHyperlinkTarget } from '../commands/hyperlinks.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import type {
  CellBatchCommand,
  CellChangeInput,
  OperationIntent,
  PermissionDecision,
} from '../commands/interaction-policy.js';
import { mergeAt } from '../commands/merge.js';
import { pivotReportEffects, pivotWrite } from '../commands/pivot-actions.js';
import { refreshPivotTable } from '../commands/pivot-table.js';
import { deferForAllowedRangeUnlock, noteChangePermitted } from '../commands/protection.js';
import { executeNoteVisibilityAction } from '../commands/ribbon-comment.js';
import { runColumnWidthDialog, runRowHeightDialog } from '../commands/row-col-dimension.js';
import { hiddenExtentInSelection } from '../commands/row-col-layout.js';
import { recordRepeatableFormatChange } from '../commands/slice-history.js';
import { rowDeleteHitsTableHeader } from '../commands/table-rows.js';
import { effectiveCellFormat } from '../engine/cell-format-projection.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { isMacHost } from '../extensions/ui-options.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { bodyBandOrigin, layoutForView, rangeRects } from '../render/geometry.js';
import { selectionRangeList } from '../store/selection-geometry.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { showDimensionDialog } from '../toolbar/dialogs/dimension.js';
import { isMenuItemDisabled, onPointerArrive, projectDisabledState } from '../toolbar/menu-a11y.js';
import { canReadClipboard, createContextMenuClipboard } from './context-menu-clipboard.js';
import {
  contextItemIds,
  contextItemToRenderEntry,
  createContextMenuItemButton,
  createContextMiniToolbarButton,
  createContextSubmenuButton,
  type RenderMenuEntry,
  sanitizeContextItems,
  selectBuiltInItems,
} from './context-menu-entries.js';
import {
  type ContextMenuFormatSortContext,
  runContextMenuFormatSortItem,
} from './context-menu-format-sort.js';
import {
  type ContextMenuInsertCopiedContext,
  runContextMenuInsertCopiedCells,
  wholeBandAxisFor,
} from './context-menu-insert-copied.js';
import type {
  ContextMenuContext,
  ContextMenuInteractionController,
  ContextMenuItem,
  ContextMenuOptions,
} from './context-menu-options.js';
import { createContextMenuPermissions } from './context-menu-permissions.js';
import { pivotMenuTarget, runContextMenuPivotItem } from './context-menu-pivot.js';
import { placeMacMenu } from './context-menu-placement.js';
import {
  buildCellEntries,
  buildColEntries,
  buildMacCellEntries,
  buildMacColEntries,
  buildMacCornerEntries,
  buildMacRowEntries,
  buildPivotEntries,
  buildRowEntries,
  compactMenuEntries,
  type DialogMenuItemId,
  type ItemId,
  isBuiltinItemId,
  isDialogMenuItemId,
  type MenuKind,
  menuEntryToContextItem,
  PASTE_QUICK_IDS,
  PLAIN_INSERT_IDS,
  toMacBandEntries,
  withoutMenuItems,
  withTableInsertMenu,
} from './context-menu-spec.js';
import {
  type ContextMenuStructureContext,
  runContextMenuStructureItem,
} from './context-menu-structure.js';
import { type MenuTarget, resolveContextMenuTarget } from './context-menu-target.js';
import { openCommentDialogFrom } from './dialog-entry.js';
import { focusSheet } from './dialog-focus.js';
import { withDialogAnchor } from './dialog-placement.js';
import { overlayPortalFor } from './overlay-portal.js';
import { panelSize, placePanel, VIEWPORT_PANEL_PAD, viewportSize } from './overlay-position.js';
import { createPivotNumberFormatPicker } from './pivot-field-settings-dialog.js';

/** Distance from a Mac submenu's top edge to the centre of its first row. */
const MAC_SUB_ROW_CENTER = 18;
/** Grace period before a submenu closes, so the pointer can cross sibling rows into it. */
const SUBMENU_CLOSE_DELAY_MS = 260;
/** How far a non-Mac submenu overlaps its parent row horizontally and sits above it. */
const SUB_OVERLAP_PX = 2;
const SUB_TOP_OFFSET_PX = 4;
/** Parking spot for a panel being measured before it is placed. */
const OFFSCREEN_PX = '-9999px';

export interface ContextMenuDeps {
  host: HTMLElement;
  /** Element whose coordinate space matches grid hit-testing. Defaults to host
   *  for standalone tests/legacy embedders. */
  grid?: HTMLElement;
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  /** UI string dictionary. Falls back to the package default (ja) if omitted. */
  strings?: Strings;
  /** Shared history. When provided, format-mutating menu actions push entries
   *  so Cmd+Z reverts them. */
  history?: History | null;
  /** Called after cut/paste/clear so caller can refresh cached cells from engine. */
  onAfterCommit?: () => void;
  /** Called when the user clicks the "Format Cells…" menu entry. */
  onFormatDialog?: () => void;
  /** Whether the dialog behind a dialog-opening item is attached; asked each time the menu
   *  opens, and an item whose dialog is not attached is hidden. Omitted: every one is. */
  dialogAttached?: (id: DialogMenuItemId) => boolean;
  /** Called when the user clicks "Paste Special…". */
  onPasteSpecial?: () => void;
  /** Called when the user clicks "Define Name…". When omitted the entry is
   *  hidden. */
  onDefineName?: () => void;
  /** Locale the Row Height / Column Width dialogs parse numbers with. */
  getLocale?: () => string;
  /** Returns the structured clipboard snapshot used by the Paste Special
   *  submenu's quick-paste actions. When it returns null those entries are
   *  disabled and only the "Paste Special…" dialog entry stays usable. */
  getClipboardSnapshot?: () => ClipboardSnapshot | null;
  /** Optional shared clipboard command path. When provided, context-menu
   *  copy/cut/paste uses it so structured snapshots and Paste Options stay
   *  consistent with keyboard shortcuts. */
  onClipboardShortcut?: (kind: 'copy' | 'cut' | 'paste') => void;
  /** Called when the user clicks "Edit comment…". When omitted the menu
   *  entry is hidden — the action requires the comment dialog feature to
   *  be wired up. */
  onEditComment?: (addr: Addr) => void;
  /** Called when the user clicks "Insert hyperlink…". When omitted the menu
   *  entry is hidden. */
  onInsertHyperlink?: () => void;
  /** Opens Pick From Drop-down List on the active cell. When omitted the menu entry is hidden. */
  onPickFromList?: () => void;
  /** Opens the active cell's hyperlink after the menu validates its target. */
  onOpenHyperlink?: (url: string) => void;
  /** Called when the user clicks the Add/Remove Watch entry. When omitted the
   *  menu entry is hidden. */
  onToggleWatch?: (addr: Addr) => void;
  /** Returns true when the active cell is currently watched. */
  isWatched?: (addr: Addr) => boolean;
  /** Called for the pivot menu's field-list entries. When omitted those entries
   *  are hidden. */
  onOpenPivotFieldList?: (sheetIndex: number, pivotIndex: number) => void;
  /** Whether the field list is showing, and closes it: the pivot menu's Hide
   *  Field List row is checked while it is hidden and toggles it. */
  isPivotFieldListOpen?: () => boolean;
  onClosePivotFieldList?: () => void;
  /** Optional composed menu contract. Omitted preserves the legacy menu. */
  options?: ContextMenuOptions;
  /** Shared authorization/execution service for restricted embeds. */
  interactionController?: ContextMenuInteractionController;
}

const itemIdSet = new Set<ItemId>([
  'bold',
  'italic',
  'underline',
  'alignLeft',
  'alignCenter',
  'alignRight',
  'borders',
  'formatCells',
]);

/** Detacher returned by `attachContextMenu`. Also exposes `setStrings` so the
 *  active dictionary can be swapped after attach. */
export interface ContextMenuHandle {
  (): void;
  /** Swap the active dictionary; takes effect on next open. */
  setStrings(next: Strings): void;
  /** Replace the menu composition at runtime; takes effect on next open. */
  setOptions(next?: ContextMenuOptions): void;
}

export function attachContextMenu(deps: ContextMenuDeps): ContextMenuHandle {
  const { host, store, wb } = deps;
  const ownerDocument = host.ownerDocument;
  const ownerWindow = ownerDocument.defaultView;
  const OwnerElement = ownerWindow?.Element ?? Element;
  const OwnerNode = ownerWindow?.Node ?? Node;
  const OwnerHTMLElement = ownerWindow?.HTMLElement ?? HTMLElement;
  const OwnerHTMLButtonElement = ownerWindow?.HTMLButtonElement ?? HTMLButtonElement;
  const hitHost = deps.grid ?? host;
  const history = deps.history ?? null;
  let options = deps.options;
  const interactionController: ContextMenuInteractionController | undefined =
    deps.interactionController ?? interactionControllerFor(store);
  if (history) wb.attachHistory(history);
  let strings = deps.strings ?? defaultStrings;
  const wrapFmt = (commandId: string, fn: () => void): void =>
    recordRepeatableFormatChange(history, store, () =>
      withSelectionFormatOrigin(store, 'contextMenu', fn, commandId),
    );
  const afterCommit = (): void => deps.onAfterCommit?.();
  const structureContext: ContextMenuStructureContext = {
    host,
    store,
    wb,
    history,
    strings: () => strings,
    afterCommit,
  };
  const formatSortContext: ContextMenuFormatSortContext = {
    store,
    wb,
    history,
    strings: () => strings,
    afterCommit,
    formatChange: wrapFmt,
  };

  const root = ownerDocument.createElement('div');
  root.className = 'fc-ctxmenu fc-ctxmenu--managed';
  root.setAttribute('role', 'menu');
  root.setAttribute('aria-label', strings.contextMenu.title);
  root.style.display = 'none';
  root.tabIndex = -1;
  const portal = overlayPortalFor(host);
  portal.appendChild(root);

  // Single reusable child panel — the context menu is one level deep.
  const sub = ownerDocument.createElement('div');
  sub.className = 'fc-ctxmenu fc-ctxmenu__sub fc-ctxmenu--managed';
  sub.setAttribute('role', 'menu');
  sub.style.display = 'none';
  sub.tabIndex = -1;
  portal.appendChild(sub);

  let visible = false;
  let menuKind: MenuKind = 'cell';
  const permissions = createContextMenuPermissions({
    store,
    interactionController,
    strings: () => strings,
  });
  const { decisionForItem, decisionReason, hasExplicitPolicy, canPasteToRange } = permissions;
  let pasteBtnRef: HTMLButtonElement | null = null;
  let activeIndex = -1;
  let focusPanel: 'root' | 'sub' = 'root';
  let restoreFocusEl: HTMLElement | null = null;
  const submenuChildren = new Map<string, RenderMenuEntry[]>();
  let openSub: { id: string; parentBtn: HTMLButtonElement } | null = null;
  let subCloseTimer: ReturnType<typeof setTimeout> | null = null;
  let activeContext: ContextMenuContext | null = null;
  let activePivot:
    | (NonNullable<ReturnType<WorkbookHandle['pivotCellInfo']>> & { sheet: number })
    | null = null;
  let activeTarget: MenuTarget | null = null;
  let activePivotAddr: Addr | null = null;
  let activeEvent: MouseEvent | KeyboardEvent | null = null;
  const activeItemSources = new Map<string, ContextMenuItem>();
  let explicitDisabledSources = new WeakSet<ContextMenuItem>();
  let policyDisabledSources = new WeakSet<ContextMenuItem>();

  const cancelSubClose = (): void => {
    if (subCloseTimer != null) {
      clearTimeout(subCloseTimer);
      subCloseTimer = null;
    }
  };

  const closeSubmenu = (): void => {
    cancelSubClose();
    if (!openSub) return;
    const parentBtn = openSub.parentBtn;
    const restoreParentFocus = visible && sub.contains(ownerDocument.activeElement);
    if (restoreParentFocus) parentBtn.focus({ preventScroll: true });
    parentBtn.setAttribute('aria-expanded', 'false');
    parentBtn.classList.remove('fc-ctxmenu__item--open');
    openSub = null;
    sub.style.display = 'none';
    clearPanelActiveItem(sub);
    sub.replaceChildren();
    if (focusPanel === 'sub' || restoreParentFocus) {
      focusPanel = 'root';
      activeIndex = panelItems(root).indexOf(parentBtn);
      setPanelActiveItem(root, parentBtn);
    }
  };

  /** Close the submenu after a short grace period — lets the pointer travel
   *  diagonally from a parent row into the child panel without it snapping
   *  shut as it crosses sibling rows. */
  const scheduleSubClose = (): void => {
    if (!openSub) return;
    cancelSubClose();
    subCloseTimer = setTimeout(() => {
      closeSubmenu();
    }, SUBMENU_CLOSE_DELAY_MS);
  };

  const hide = (restoreFocus = false): void => {
    if (!visible) return;
    visible = false;
    closeSubmenu();
    clearActiveItems();
    root.style.display = 'none';
    activeIndex = -1;
    focusPanel = 'root';
    const focusTarget = restoreFocusEl;
    restoreFocusEl = null;
    if (restoreFocus) {
      if (focusTarget) focusTarget.focus({ preventScroll: true });
      else focusSheet(host, { preventScroll: true });
    }
  };

  /** `row` is the clicked item: a dialog the action opens centres on it. The anchor lasts only for this dispatch. */
  const invokeMenuAction = (action: () => void, row?: HTMLElement): void => {
    const rowRect = row?.getBoundingClientRect();
    const anchor =
      rowRect && rowRect.width > 0
        ? { x: rowRect.left + rowRect.width / 2, y: rowRect.top + rowRect.height / 2 }
        : null;
    const menuHadFocus =
      root.contains(ownerDocument.activeElement) || sub.contains(ownerDocument.activeElement);
    hide(false);
    // Focus the host before the action so keyboard undo works and an opened dialog keeps focus.
    if (menuHadFocus) focusSheet(host, { preventScroll: true });
    withDialogAnchor(anchor, action);
  };

  const isMacPlatform = (): boolean => isMacHost(host);

  /** Disable a button on behalf of the host policy, remembering that the policy did it. */
  const markPolicyDisabled = (
    btn: HTMLButtonElement,
    decision: ReturnType<typeof decisionForItem>,
  ): void => {
    btn.dataset.policyDisabled = 'true';
    setContextMenuItemDisabled(btn, true, decisionReason(decision));
  };

  const activePanel = (): HTMLElement => (focusPanel === 'sub' ? sub : root);

  const panelItems = (panel: HTMLElement): HTMLButtonElement[] =>
    Array.from(panel.querySelectorAll<HTMLButtonElement>('.fc-ctxmenu__item')).filter(
      (btn) => !isMenuItemDisabled(btn),
    );

  const clearPanelActiveItem = (panel: HTMLElement): void => {
    for (const item of panel.querySelectorAll('.fc-ctxmenu__item--active')) {
      item.classList.remove('fc-ctxmenu__item--active');
    }
  };

  const setPanelActiveItem = (panel: HTMLElement, item: HTMLButtonElement | null): void => {
    clearPanelActiveItem(panel);
    if (item && !isMenuItemDisabled(item)) {
      item.classList.add('fc-ctxmenu__item--active');
    }
  };

  const clearActiveItems = (): void => {
    clearPanelActiveItem(root);
    clearPanelActiveItem(sub);
  };

  const createContext = (
    target: MenuTarget,
    event: MouseEvent | KeyboardEvent,
    defaultItems: readonly ContextMenuItem[],
  ): ContextMenuContext => {
    const selection = { ...store.getState().selection.range };
    const canExecute = (item: ItemId | ContextMenuItem | OperationIntent): PermissionDecision =>
      decisionForItem(item);
    return {
      kind: target.kind,
      cell: { ...target.cell },
      selection,
      defaultItems,
      event,
      canExecute,
      permission: canExecute,
    };
  };

  const decorateContextItems = (items: readonly ContextMenuItem[]): ContextMenuItem[] =>
    items.map((item) => {
      const children = item.children ? decorateContextItems(item.children) : undefined;
      if (item.disabled) {
        const decorated = children ? { ...item, children } : item;
        if (explicitDisabledSources.has(item)) explicitDisabledSources.add(decorated);
        return decorated;
      }
      const decision = item.builtIn || item.command ? decisionForItem(item) : null;
      if (decision && !decision.allowed) {
        const decorated = {
          ...item,
          ...(children ? { children } : {}),
          disabled: true,
          disabledReason: decisionReason(decision) ?? undefined,
        };
        policyDisabledSources.add(decorated);
        if (explicitDisabledSources.has(item)) explicitDisabledSources.add(decorated);
        return decorated;
      }
      return children ? { ...item, children } : item;
    });

  const markExplicitDisabledSources = (items: readonly ContextMenuItem[]): void => {
    for (const item of items) {
      if (item.disabled && !policyDisabledSources.has(item)) explicitDisabledSources.add(item);
      if (item.children) markExplicitDisabledSources(item.children);
    }
  };

  const focusMenuItem = (idx: number): void => {
    const panelName = focusPanel;
    const panel = activePanel();
    const items = panelItems(panel);
    if (items.length === 0) {
      activeIndex = -1;
      setPanelActiveItem(panel, null);
      return;
    }
    const nextIndex = ((idx % items.length) + items.length) % items.length;
    const next = items[nextIndex];
    if (panelName === 'root' && openSub && openSub.parentBtn !== next) closeSubmenu();
    activeIndex = nextIndex;
    setPanelActiveItem(panel, next ?? null);
    next?.focus();
  };

  const followPointer = (panel: 'root' | 'sub', btn: HTMLButtonElement): void => {
    const menuPanel = panel === 'root' ? root : sub;
    if (isMenuItemDisabled(btn)) {
      focusPanel = panel;
      activeIndex = -1;
      setPanelActiveItem(menuPanel, null);
      menuPanel.focus({ preventScroll: true });
      return;
    }
    focusPanel = panel;
    activeIndex = panelItems(menuPanel).indexOf(btn);
    if (activeIndex < 0) {
      setPanelActiveItem(menuPanel, null);
      return;
    }
    setPanelActiveItem(menuPanel, btn);
    if (panel === 'sub' && openSub) setPanelActiveItem(root, openSub.parentBtn);
    btn.focus({ preventScroll: true });
  };

  const openSubmenu = (id: string, parentBtn: HTMLButtonElement): void => {
    cancelSubClose();
    if (openSub?.id === id) return;
    closeSubmenu();
    const children = submenuChildren.get(id);
    if (!children) return;
    const disabled = new Set<string>();
    if (id === 'pasteSpecialMenu' && !clipboard.snapshot()) {
      for (const d of PASTE_QUICK_IDS) disabled.add(d);
    }
    sub.replaceChildren();
    sub.dataset.fcSubmenu = id;
    for (const child of children) appendEntry('sub', sub, child, disabled);
    sub.style.display = 'block';
    sub.style.left = OFFSCREEN_PX;
    sub.style.top = OFFSCREEN_PX;
    const r = parentBtn.getBoundingClientRect();
    const { width: sw } = panelSize(sub);
    const viewport = viewportSize(ownerDocument);
    // The desktop submenu overlaps its parent by the panel padding and centres its
    // first row on the parent row.
    const mac = isMacPlatform();
    let x = mac ? r.right : r.right - SUB_OVERLAP_PX;
    if (x + sw > viewport.width - VIEWPORT_PANEL_PAD) {
      x = mac ? r.left - sw : r.left - sw + SUB_OVERLAP_PX;
    }
    placePanel(sub, x, mac ? r.top + r.height / 2 - MAC_SUB_ROW_CENTER : r.top - SUB_TOP_OFFSET_PX);
    openSub = { id, parentBtn };
    parentBtn.setAttribute('aria-expanded', 'true');
    parentBtn.classList.add('fc-ctxmenu__item--open');
    setPanelActiveItem(root, parentBtn);
  };

  const appendEntry = (
    panel: 'root' | 'sub',
    container: HTMLElement,
    entry: RenderMenuEntry,
    disabledIds: Set<string>,
  ): void => {
    if (entry.kind === 'sep') {
      const sep = container.ownerDocument.createElement('hr');
      sep.className = 'fc-ctxmenu__sep';
      container.appendChild(sep);
      return;
    }
    if (entry.kind === 'submenu') {
      submenuChildren.set(entry.id, entry.children);
      const source = entry.source;
      if (source) activeItemSources.set(entry.id, source);
      const btn = createContextSubmenuButton(entry, ownerDocument);
      if (source && explicitDisabledSources.has(source)) {
        setContextMenuItemDisabled(btn, true, source.disabledReason ?? null);
      } else if ((source?.builtIn || source?.command) && !decisionForItem(source).allowed) {
        const decision = decisionForItem(source);
        btn.dataset.policyDisabled = 'true';
        setContextMenuItemDisabled(btn, true, decisionReason(decision));
      }
      onPointerArrive(btn, () => {
        if (isMenuItemDisabled(btn)) {
          if (panel === 'root' && openSub && openSub.parentBtn !== btn) scheduleSubClose();
          followPointer(panel, btn);
          return;
        }
        openSubmenu(entry.id, btn);
        followPointer(panel, btn);
      });
      btn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (btn.disabled) return;
        openSubmenu(entry.id, btn);
        focusPanel = 'sub';
        focusMenuItem(0);
      });
      container.appendChild(btn);
      return;
    }
    const btn = createContextMenuItemButton(entry, ownerDocument);
    const source = entry.source;
    if (source) activeItemSources.set(entry.id, source);
    if (disabledIds.has(entry.id)) {
      setContextMenuItemDisabled(btn, true, strings.contextMenu.pasteSpecialRequiresCopiedCells);
    } else if (source && explicitDisabledSources.has(source)) {
      if ((source.builtIn || source.command) && !decisionForItem(source).allowed) {
        btn.dataset.policyDisabled = 'true';
      }
      setContextMenuItemDisabled(btn, true, source.disabledReason ?? null);
    } else if (source?.builtIn || source?.command) {
      const decision = decisionForItem(source);
      if (!decision.allowed) markPolicyDisabled(btn, decision);
    }
    btn.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      if (btn.disabled) return;
      invokeMenuAction(() => runEntry(entry), btn);
    });
    onPointerArrive(btn, () => {
      if (panel === 'root') {
        if (openSub && openSub.parentBtn !== btn) scheduleSubClose();
      } else if (!isMenuItemDisabled(btn)) {
        cancelSubClose();
      }
      followPointer(panel, btn);
    });
    container.appendChild(btn);
    if (panel === 'root' && entry.id === 'paste') pasteBtnRef = btn;
  };

  const buildMiniToolbar = (
    allowedIds?: ReadonlySet<string>,
    hiddenIds?: ReadonlySet<string>,
  ): HTMLElement => {
    const state = store.getState();
    const activeAddr = state.selection.active;
    const merge = mergeAt(state, activeAddr);
    const formatAddr = merge ? { sheet: merge.sheet, row: merge.r0, col: merge.c0 } : activeAddr;
    const format = effectiveCellFormat(wb, state, formatAddr);
    const pressedById: Partial<Record<ItemId, boolean>> = {
      bold: format?.bold === true,
      italic: format?.italic === true,
      underline: Boolean(format?.underline),
      alignLeft: format?.align === 'left',
      alignCenter: format?.align === 'center',
      alignRight: format?.align === 'right',
    };
    const toolbar = ownerDocument.createElement('div');
    toolbar.className = 'fc-ctxmenu__mini';
    toolbar.setAttribute('role', 'toolbar');
    toolbar.setAttribute('aria-label', strings.contextMenu.title);

    const buttons: readonly { id: ItemId; label: string }[] = [
      { id: 'bold', label: strings.contextMenu.bold },
      { id: 'italic', label: strings.contextMenu.italic },
      { id: 'underline', label: strings.contextMenu.underline },
      { id: 'alignLeft', label: strings.contextMenu.alignLeft },
      { id: 'alignCenter', label: strings.contextMenu.alignCenter },
      { id: 'alignRight', label: strings.contextMenu.alignRight },
      { id: 'borders', label: strings.contextMenu.borders },
      { id: 'formatCells', label: strings.contextMenu.formatCells },
    ];

    for (const item of buttons) {
      if (allowedIds && !allowedIds.has(item.id)) continue;
      if (hiddenIds?.has(item.id)) continue;
      const btn = createContextMiniToolbarButton(item, ownerDocument);
      const pressed = pressedById[item.id];
      if (pressed !== undefined) btn.setAttribute('aria-pressed', String(pressed));
      const decision = decisionForItem(item.id);
      if (!decision.allowed) markPolicyDisabled(btn, decision);
      btn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        invokeMenuAction(() => run(item.id), btn);
      });
      toolbar.appendChild(btn);
    }

    return toolbar;
  };

  const buildMenu = (kind: MenuKind, corner = false): void => {
    if (openSub) closeSubmenu();
    clearActiveItems();
    activeIndex = -1;
    menuKind = kind;
    root.replaceChildren();
    submenuChildren.clear();
    activeItemSources.clear();
    explicitDisabledSources = new WeakSet<ContextMenuItem>();
    policyDisabledSources = new WeakSet<ContextMenuItem>();
    pasteBtnRef = null;
    const mac = isMacPlatform();
    const activeAddr = store.getState().selection.active;
    const hasNote = commentAt(store.getState(), activeAddr) !== null;
    // The report cell under the pointer, which a right-click does not always make active.
    const pivotAddr = kind === 'cell' && activeTarget ? activeTarget.cell : activeAddr;
    const pivot = kind === 'cell' ? wb.pivotCellInfo(pivotAddr) : null;
    activePivot = pivot ? { ...pivot, sheet: pivotAddr.sheet } : null;
    activePivotAddr = pivot ? { ...pivotAddr } : null;
    // Menus measured apart may differ in metrics; the stylesheet tells them apart by this.
    root.dataset.fcMenu = pivot ? 'pivot' : kind;
    sub.dataset.fcMenu = root.dataset.fcMenu;
    const raw = pivot
      ? buildPivotEntries(
          strings,
          pivotMenuTarget(
            wb,
            { addr: pivotAddr, info: pivot },
            !!deps.onOpenPivotFieldList,
            deps.isPivotFieldListOpen ? !deps.isPivotFieldListOpen() : undefined,
          ),
        )
      : kind === 'row'
        ? mac
          ? buildMacRowEntries(strings)
          : buildRowEntries(strings)
        : kind === 'col'
          ? mac
            ? buildMacColEntries(strings)
            : buildColEntries(strings)
          : mac
            ? (corner ? buildMacCornerEntries : buildMacCellEntries)(strings, hasNote)
            : buildCellEntries(strings, hasNote);
    const stateNow = store.getState();
    const bandRange = stateNow.selection.range;
    const inTable =
      kind === 'cell' && !pivot && !stateNow.ui.copyRange
        ? tableForCell(stateNow.tables.tables, activeAddr.sheet, activeAddr.row, activeAddr.col)
        : null;
    const baseEntries = mac && pivot ? toMacBandEntries(raw, strings) : raw;
    const platformEntries = inTable
      ? withTableInsertMenu(baseEntries, strings, {
          rows: !(inTable.showHeader && bandRange.r0 === inTable.range.r0),
        })
      : baseEntries;
    const hasCopiedCells = !!store.getState().ui.copyRange;
    const pendingSnapshot = hasCopiedCells ? clipboard.snapshot() : null;
    const pendingCutBand =
      pendingSnapshot?.mode === 'cut' && wholeBandAxisFor(pendingSnapshot) === kind;
    const watched = !!deps.isWatched?.(activeAddr);
    const entries = compactMenuEntries(
      withoutMenuItems(
        platformEntries,
        (id) => isDialogMenuItemId(id) && deps.dialogAttached?.(id) === false,
      )
        .filter(
          (e) => !(e.kind === 'item' && e.id === 'insertHyperlink' && !deps.onInsertHyperlink),
        )
        .filter((e) => !(e.kind === 'item' && e.id === 'openHyperlink' && !deps.onOpenHyperlink))
        // A pending copy swaps the plain insert entries for "Insert Copied
        // Cells" — the desktop menus never show both at once.
        .filter((e) => !(e.kind === 'item' && e.id === 'insertCopiedCells' && !hasCopiedCells))
        .filter((e) => !(e.kind === 'item' && hasCopiedCells && PLAIN_INSERT_IDS.includes(e.id)))
        .filter((e) => !(e.kind === 'item' && e.id === 'insertComment' && !deps.onEditComment))
        .filter((e) => !(e.kind === 'item' && e.id === 'pickFromList' && !deps.onPickFromList))
        .filter(
          (e) => !(e.kind === 'item' && e.id === 'newComment' && !wb.capabilities.threadedComments),
        )
        .filter((e) => !(e.kind === 'item' && e.id === 'showHideNote' && !hasNote))
        // A protected sheet without Edit objects drops a note's edit and delete rows entirely,
        // and on the Mac New Note as well.
        .filter(
          (e) =>
            !(
              e.kind === 'item' &&
              (e.id === 'insertComment' || e.id === 'deleteComment') &&
              (mac || hasNote) &&
              !noteChangePermitted(store.getState(), activeAddr)
            ),
        )
        .filter((e) => !(e.kind === 'item' && e.id === 'toggleWatch' && !deps.onToggleWatch))
        .filter((e) => !(e.kind === 'item' && e.id === 'defineName' && !deps.onDefineName))
        .map((e) => {
          if (e.kind === 'item' && e.id === 'toggleWatch') {
            return {
              ...e,
              label: watched ? strings.contextMenu.removeWatch : strings.contextMenu.addWatch,
            };
          }
          // No delete is offered over a table's header row.
          if (
            e.kind === 'item' &&
            e.id === 'rowDelete' &&
            rowDeleteHitsTableHeader(
              store,
              bandRange.sheet,
              bandRange.r0,
              bandRange.r1 - bandRange.r0 + 1,
            )
          ) {
            return { ...e, disabled: true };
          }
          if (e.kind === 'item' && e.id === 'insertCopiedCells' && pendingCutBand) {
            return { ...e, label: strings.contextMenu.insertCutCells };
          }
          return e;
        }),
    );
    const defaultItems = entries.map(menuEntryToContextItem);
    let visibleItems = decorateContextItems(defaultItems);
    const builtInOptions = options?.mode === 'builtIn' ? options : undefined;
    if (builtInOptions?.items) {
      visibleItems = selectBuiltInItems(visibleItems, new Set(builtInOptions.items));
    }
    let composedContext = activeContext;
    if (builtInOptions?.transform) {
      if (activeTarget && activeEvent) {
        composedContext = createContext(activeTarget, activeEvent, visibleItems);
      }
      if (composedContext) {
        try {
          visibleItems = sanitizeContextItems(builtInOptions.transform(composedContext));
        } catch (err) {
          console.warn('formulon-cell: context-menu transform failed', err);
        }
      }
    } else {
      visibleItems = sanitizeContextItems(visibleItems);
    }
    markExplicitDisabledSources(visibleItems);
    if (composedContext) activeContext = composedContext;
    const renderEntries = visibleItems.map(contextItemToRenderEntry);
    const visibleRootIds = contextItemIds(visibleItems);
    // The Mac menu carries no mini formatting toolbar.
    if (kind === 'cell' && !mac && !pivot) {
      const allowedMini = builtInOptions?.items
        ? new Set<string>(builtInOptions.items)
        : builtInOptions
          ? itemIdSet
          : undefined;
      root.appendChild(buildMiniToolbar(allowedMini, options ? visibleRootIds : undefined));
    }
    const noDisabled = new Set<string>();
    for (const entry of renderEntries) appendEntry('root', root, entry, noDisabled);

    const s = store.getState();
    const rowHidden =
      kind === 'row' &&
      hiddenExtentInSelection(s.layout, 'row', s.selection.range.r0, s.selection.range.r1) !== null;
    const colHidden =
      kind === 'col' &&
      hiddenExtentInSelection(s.layout, 'col', s.selection.range.c0, s.selection.range.c1) !== null;
    const rowUnhide = root.querySelector<HTMLButtonElement>('[data-fc-action="rowUnhide"]');
    const colUnhide = root.querySelector<HTMLButtonElement>('[data-fc-action="colUnhide"]');
    const openHyperlink = root.querySelector<HTMLButtonElement>('[data-fc-action="openHyperlink"]');
    // The Mac desktop menu leaves Unhide enabled whatever the selection holds.
    if (rowUnhide && !mac) {
      setContextMenuItemDisabled(rowUnhide, !rowHidden, strings.contextMenu.noHiddenRows);
    }
    if (colUnhide && !mac) {
      setContextMenuItemDisabled(colUnhide, !colHidden, strings.contextMenu.noHiddenColumns);
    }
    if (openHyperlink) {
      const target = hyperlinkAt(s, s.selection.active);
      setContextMenuItemDisabled(
        openHyperlink,
        target === null || !isFollowableHyperlinkTarget(target),
        strings.ribbonMenu.linkNoHyperlink,
      );
    }
  };

  const show = (
    clientX: number,
    clientY: number,
    target: MenuTarget,
    event: MouseEvent | KeyboardEvent,
  ): void => {
    if (options?.mode === 'disabled') return;
    if (visible) hide(false);
    // A fullscreen/dialog resolver can change between the browser's
    // fullscreenchange notification and this key/mouse event. Refresh the
    // owned portal synchronously so the menu is born inside the current focus
    // boundary instead of one event behind.
    overlayPortalFor(host);
    activeTarget = target;
    activeEvent = event;
    activeContext = options ? createContext(target, event, []) : null;
    const currentActive = ownerDocument.activeElement;
    restoreFocusEl =
      currentActive instanceof OwnerHTMLElement &&
      !root.contains(currentActive) &&
      !sub.contains(currentActive)
        ? currentActive
        : host;
    root.setAttribute('aria-label', strings.contextMenu.title);
    focusPanel = 'root';
    buildMenu(target.kind, target.corner === true);
    if (options?.mode === 'host') {
      const context = activeContext;
      root.style.display = 'none';
      activeContext = null;
      if (context) options.onOpen(context);
      return;
    }
    if (pasteBtnRef) {
      const canPaste = canReadClipboard();
      if (!canPaste) {
        setContextMenuItemDisabled(pasteBtnRef, true, strings.contextMenu.clipboardUnavailable);
      } else {
        const decision = decisionForItem('paste');
        if (!decision.allowed) {
          setContextMenuItemDisabled(pasteBtnRef, true, decisionReason(decision));
        } else if (!hasExplicitPolicy()) {
          // Preserve the legacy accessibility state: old callers observed an
          // explicit aria-disabled="false" on an available paste action.
          setContextMenuItemDisabled(pasteBtnRef, false, null);
        }
      }
    }
    root.style.display = 'block';
    root.style.left = OFFSCREEN_PX;
    root.style.top = OFFSCREEN_PX;
    visible = true;
    const mac = isMacPlatform();
    if (mac) placeMacMenu(root, clientX, clientY);
    else placePanel(root, clientX, clientY);
    // A Mac menu opened by the pointer starts with no row highlighted; the
    // first arrow key then lands on the first row.
    if (mac && event.type === 'contextmenu') {
      activeIndex = -1;
      root.focus({ preventScroll: true });
    } else {
      focusMenuItem(0);
    }
  };

  const isOwnChromeContextTarget = (target: EventTarget | null): boolean =>
    target instanceof OwnerElement &&
    !!target.closest(
      '.fc-host__formulabar, .fc-host__sheetbar, .fc-host__statusbar, .fc-sheetmenu, .fc-statusbar__chooser',
    );

  const onContextMenu = (e: MouseEvent): void => {
    if (isOwnChromeContextTarget(e.target)) return;
    if (options?.mode === 'disabled') return;
    e.preventDefault();
    const target = resolveContextMenuTarget(
      store,
      hitHost,
      e,
      options !== undefined,
      permissions.canChangeSelection,
    );
    show(e.clientX, e.clientY, target, e);
  };

  const onHostKey = (e: KeyboardEvent): void => {
    if (options?.mode === 'disabled') return;
    if (e.key !== 'ContextMenu' && !(e.key === 'F10' && e.shiftKey)) return;
    e.preventDefault();
    const active = store.getState().selection.active;
    const { x, y } = keyboardMenuPoint();
    show(
      x,
      y,
      {
        kind: 'cell',
        cell: { ...active },
      },
      e,
    );
  };

  /** Position keyboard-opened menus from the active cell's visible geometry.
   *  The cell layout is local to the grid; the host rect converts it to client
   *  coordinates before the shared overlay positioning handles clamping and
   *  any transformed portal. */
  const keyboardMenuPoint = (): { x: number; y: number } => {
    const state = store.getState();
    const layout = layoutForView(state);
    const viewport = state.viewport;
    const gridRect = hitHost.getBoundingClientRect();
    const fallback = (): { x: number; y: number } => {
      const origin = bodyBandOrigin(layout, viewport);
      return { x: gridRect.left + origin.x, y: gridRect.top + origin.y };
    };

    if (
      !Number.isFinite(gridRect.left) ||
      !Number.isFinite(gridRect.top) ||
      gridRect.width <= 0 ||
      gridRect.height <= 0
    ) {
      return fallback();
    }

    const active = state.selection.active;
    if (active.sheet !== state.data.sheetIndex) return fallback();

    const activeRange = {
      r0: active.row,
      r1: active.row,
      c0: active.col,
      c1: active.col,
    };
    const merge = mergeAt(state, active);
    const fragments = rangeRects(layout, viewport, merge ?? activeRange).filter(
      (rect) => rect.w > 0 && rect.h > 0,
    );
    if (fragments.length === 0) return fallback();

    // `rangeRects` clips to the viewport, so this stays bounded for a far-away row.
    const activeCell = merge
      ? rangeRects(layout, viewport, activeRange).find((rect) => rect.w > 0 && rect.h > 0)
      : fragments[0];

    const fragment = activeCell
      ? fragments.find(
          (rect) =>
            activeCell.x + activeCell.w / 2 >= rect.x &&
            activeCell.x + activeCell.w / 2 <= rect.x + rect.w &&
            activeCell.y + activeCell.h / 2 >= rect.y &&
            activeCell.y + activeCell.h / 2 <= rect.y + rect.h,
        )
      : undefined;
    const visibleFragment = fragment ?? fragments[0];
    if (!visibleFragment) return fallback();

    return {
      x: gridRect.left + visibleFragment.x + (layout.rtl ? visibleFragment.w : 0),
      y: gridRect.top + visibleFragment.y + visibleFragment.h,
    };
  };

  const insideMenu = (target: EventTarget | null): boolean =>
    target instanceof OwnerNode && (root.contains(target) || sub.contains(target));

  const onDocFocusIn = (event: FocusEvent): void => {
    if (!visible) return;
    const target = event.target;
    if (!insideMenu(target)) {
      hide(false);
      return;
    }
    if (target === root || target === sub) {
      const panel = target === root ? root : sub;
      focusPanel = target === root ? 'root' : 'sub';
      activeIndex = -1;
      setPanelActiveItem(panel, null);
      if (target === sub && openSub) setPanelActiveItem(root, openSub.parentBtn);
      return;
    }
    if (
      !(target instanceof OwnerHTMLButtonElement) ||
      !target.classList.contains('fc-ctxmenu__item')
    )
      return;
    const panelName = root.contains(target) ? 'root' : 'sub';
    const panel = panelName === 'root' ? root : sub;
    focusPanel = panelName;
    if (isMenuItemDisabled(target)) {
      activeIndex = -1;
      setPanelActiveItem(panel, null);
    } else {
      activeIndex = panelItems(panel).indexOf(target);
      setPanelActiveItem(panel, target);
    }
    if (panelName === 'sub' && openSub) setPanelActiveItem(root, openSub.parentBtn);
  };

  const onDocOutsidePress = (e: MouseEvent): void => {
    if (!visible) return;
    if (insideMenu(e.target)) return;
    hide(false);
  };

  const swallow = (e: Event): void => {
    e.preventDefault();
    e.stopPropagation();
  };

  const closeSubmenuToParent = (): void => {
    if (!openSub) return;
    const parent = openSub.parentBtn;
    closeSubmenu();
    focusPanel = 'root';
    parent.focus();
    activeIndex = panelItems(root).indexOf(parent);
  };

  const onDocKey = (e: KeyboardEvent): void => {
    if (!visible) return;
    if (!insideMenu(ownerDocument.activeElement)) {
      hide(false);
      return;
    }
    if (e.key === 'Tab') {
      hide(false);
      return;
    }
    if (e.key === 'Escape') {
      swallow(e);
      if (openSub) closeSubmenuToParent();
      else hide(true);
    } else if (e.key === 'ArrowDown') {
      swallow(e);
      focusMenuItem(activeIndex + 1);
    } else if (e.key === 'ArrowUp') {
      swallow(e);
      focusMenuItem(activeIndex - 1);
    } else if (e.key === 'Home') {
      swallow(e);
      focusMenuItem(0);
    } else if (e.key === 'End') {
      swallow(e);
      focusMenuItem(panelItems(activePanel()).length - 1);
    } else if (e.key === 'ArrowRight') {
      const el = ownerDocument.activeElement;
      if (el instanceof OwnerHTMLButtonElement && el.dataset.fcSubmenu) {
        swallow(e);
        openSubmenu(el.dataset.fcSubmenu, el);
        focusPanel = 'sub';
        focusMenuItem(0);
      }
    } else if (e.key === 'ArrowLeft') {
      if (openSub && focusPanel === 'sub') {
        swallow(e);
        closeSubmenuToParent();
      }
    } else if (e.key === 'Enter' || e.key === ' ') {
      const target = ownerDocument.activeElement;
      if (target instanceof OwnerHTMLButtonElement && insideMenu(target)) {
        swallow(e);
        target.click();
      }
    }
  };

  const onScroll = (event: Event): void => {
    if (insideMenu(event.target)) return;
    hide(false);
  };

  const onResize = (): void => {
    if (!visible) return;
    hide(insideMenu(ownerDocument.activeElement));
  };

  const refreshOpenPermissions = (): void => {
    if (!visible || !interactionController) return;
    const panels = [root, sub];
    for (const panel of panels) {
      for (const button of panel.querySelectorAll<HTMLButtonElement>(
        '.fc-ctxmenu__item, .fc-ctxmenu__mini-btn',
      )) {
        const id = button.dataset.fcAction;
        const source = id ? activeItemSources.get(id) : undefined;
        if (!id || (!source && !isBuiltinItemId(id))) continue;
        const decision = source ? decisionForItem(source) : decisionForItem(id as ItemId);
        if (!decision.allowed) {
          markPolicyDisabled(button, decision);
        } else if (button.dataset.policyDisabled === 'true') {
          delete button.dataset.policyDisabled;
          if (source && explicitDisabledSources.has(source)) {
            setContextMenuItemDisabled(button, true, source.disabledReason ?? null);
          } else {
            setContextMenuItemDisabled(button, false, null);
          }
        }
      }
    }
  };

  root.addEventListener('mouseleave', () => {
    if (!visible || openSub || !isMacPlatform()) return;
    activeIndex = -1;
    root.focus({ preventScroll: true });
  });
  onPointerArrive(sub, () => {
    cancelSubClose();
    if (!openSub) return;
    setPanelActiveItem(root, openSub.parentBtn);
    focusPanel = 'sub';
    const active = sub.querySelector<HTMLButtonElement>('.fc-ctxmenu__item--active');
    if (active && panelItems(sub).includes(active)) {
      activeIndex = panelItems(sub).indexOf(active);
      active.focus({ preventScroll: true });
    } else {
      activeIndex = -1;
      setPanelActiveItem(sub, null);
      sub.focus({ preventScroll: true });
    }
  });
  sub.addEventListener('mouseleave', scheduleSubClose);

  function runEntry(entry: RenderMenuEntry): void {
    const source = entry.source;
    if (source?.disabled) return;
    if (source && !decisionForItem(source).allowed) return;
    if (source?.builtIn) {
      run(source.builtIn);
      return;
    }
    if (source?.command) {
      // Commands supplied by a host are still dispatched through the
      // context-menu route. Never trust a caller-provided origin here: an
      // arbitrary origin could bypass the route's policy decision between
      // rendering and activation.
      executeBatch({ ...source.command, origin: 'contextMenu' });
      return;
    }
    if (source?.action) {
      if (activeContext) source.action(activeContext);
      return;
    }
    // A source-less entry is legacy only. Unknown IDs must never fall through
    // the built-in action switch.
    if (!source && isBuiltinItemId(entry.id)) run(entry.id);
  }

  const executeBatch = (command: CellBatchCommand): void => {
    if (!interactionController) return;
    try {
      const result = interactionController.execute(command);
      if (result && typeof result === 'object' && 'then' in result) {
        void Promise.resolve(result).then(() => deps.onAfterCommit?.());
      } else {
        deps.onAfterCommit?.();
      }
    } catch (err) {
      console.warn('formulon-cell: context-menu batch failed', err);
    }
  };

  const executeClipboardText = (text: string): void => {
    const state = store.getState();
    const start = state.selection.range;
    const rows = parseTSV(text);
    const changes: CellChangeInput[] = [];
    rows.forEach((row, rowOffset) => {
      row.forEach((input, colOffset) => {
        changes.push({
          addr: {
            sheet: start.sheet,
            row: start.r0 + rowOffset,
            col: start.c0 + colOffset,
          },
          input,
        });
      });
    });
    if (changes.length === 0) return;
    executeBatch({
      type: 'cellBatch',
      operation: 'paste',
      origin: 'contextMenu',
      changes,
    });
  };

  const clipboard = createContextMenuClipboard({
    store,
    wb,
    history,
    deps,
    afterCommit,
    hasExplicitPolicy,
    canPasteToRange,
    executeClipboardText,
  });

  const insertCopiedContext: ContextMenuInsertCopiedContext = {
    host,
    store,
    wb,
    history,
    strings: () => strings,
    afterCommit,
    clipboard,
  };

  const refreshPivotCells = (sheet: number): void => {
    mutators.replaceCells(store, wb.cells(sheet));
    deps.onAfterCommit?.();
  };

  const pivotNumberFormat = createPivotNumberFormatPicker({
    host,
    store,
    getStrings: () => strings,
    getWb: () => wb,
    ...(deps.getLocale ? { getLocale: deps.getLocale } : {}),
  });
  function run(id: string): void {
    if (!isBuiltinItemId(id)) return;
    if (!decisionForItem(id).allowed) return;
    if (permissions.navigationBlocks(id)) return;
    const state = store.getState();
    if (activePivot && id !== 'pivotRefresh') {
      const ref = { addr: activePivotAddr ?? state.selection.active, info: activePivot };
      const pivotContext = {
        host,
        store,
        wb,
        strings,
        history,
        pickNumberFormat: pivotNumberFormat.pick,
        refresh: refreshPivotCells,
        ...(activeEvent && 'clientX' in activeEvent
          ? { anchor: { x: activeEvent.clientX, y: activeEvent.clientY } }
          : {}),
        ...(deps.onOpenPivotFieldList ? { onOpenPivotFieldList: deps.onOpenPivotFieldList } : {}),
        ...(deps.isPivotFieldListOpen ? { isPivotFieldListOpen: deps.isPivotFieldListOpen } : {}),
        ...(deps.onClosePivotFieldList
          ? { onClosePivotFieldList: deps.onClosePivotFieldList }
          : {}),
      };
      if (runContextMenuPivotItem(pivotContext, ref, id)) return;
    }
    if (clipboard.run(id, state)) return;
    if (runContextMenuStructureItem(structureContext, id, state)) return;
    if (runContextMenuFormatSortItem(formatSortContext, id, state)) return;
    switch (id) {
      case 'pasteSpecial': {
        deps.onPasteSpecial?.();
        return;
      }
      case 'insertCopiedCells': {
        runContextMenuInsertCopiedCells(insertCopiedContext, menuKind);
        return;
      }
      case 'clear': {
        const selected = state.selection;
        if (deferForAllowedRangeUnlock(store, selectionRangeList(selected), () => run(id))) {
          return;
        }
        if (hasExplicitPolicy()) {
          const addresses = collectSelectedContentAddresses(store, wb);
          if (addresses.length === 0) return;
          executeBatch({
            type: 'cellBatch',
            operation: 'clear',
            origin: 'contextMenu',
            commandId: 'clear',
            changes: addresses.map((addr) => ({ addr, value: { kind: 'blank' as const } })),
          });
          return;
        }
        try {
          const result = clearSelectedContents({
            store,
            workbook: wb,
            history,
            origin: 'contextMenu',
            commandId: 'clear',
          });
          if (result.status === 'applied' && result.applied.length > 0) deps.onAfterCommit?.();
        } catch (err) {
          console.warn('formulon-cell: context-menu clear failed', err);
        }
        return;
      }
      case 'formatCells': {
        deps.onFormatDialog?.();
        return;
      }
      case 'defineName': {
        deps.onDefineName?.();
        return;
      }
      case 'rowHeight':
      case 'colWidth': {
        const rows = id === 'rowHeight';
        const runDialog = rows ? runRowHeightDialog : runColumnWidthDialog;
        void runDialog({
          store,
          history,
          wb,
          range: state.selection.range,
          origin: 'contextMenu',
          commandId: id,
          prompt: (initial, max) =>
            showDimensionDialog({
              axis: rows ? 'row' : 'col',
              strings,
              initial,
              max,
              locale: deps.getLocale,
            }),
        });
        return;
      }
      case 'selectAll': {
        mutators.selectAll(store);
        return;
      }
      case 'insertComment': {
        deps.onEditComment?.(state.selection.active);
        return;
      }
      case 'deleteComment': {
        const addr = state.selection.active;
        const ctx = { store, history, origin: 'contextMenu', commandId: 'deleteComment' } as const;
        commentWrite(ctx, [addr], undefined, () =>
          recordCommentChange(history, store, wb, [addr], () => clearComment(store, addr, wb), ctx),
        );
        return;
      }
      case 'showHideNote': {
        executeNoteVisibilityAction({ store, workbook: wb, action: 'toggle-active' });
        return;
      }
      case 'insertHyperlink': {
        deps.onInsertHyperlink?.();
        return;
      }
      case 'newComment': {
        openCommentDialogFrom(store, 'thread', { origin: 'contextMenu', commandId: 'newComment' });
        return;
      }
      case 'pickFromList': {
        deps.onPickFromList?.();
        return;
      }
      case 'openHyperlink': {
        const target = hyperlinkAt(state, state.selection.active);
        if (!target || !isFollowableHyperlinkTarget(target)) return;
        deps.onOpenHyperlink?.(target);
        return;
      }
      case 'toggleWatch': {
        deps.onToggleWatch?.(state.selection.active);
        return;
      }
      case 'pivotRefresh': {
        if (!activePivot) return;
        const pivot = activePivot;
        const result = pivotWrite(
          { store, origin: 'contextMenu', commandId: 'pivotRefresh' },
          history,
          pivotReportEffects(wb, pivot.sheet, pivot.pivotIndex),
          { ok: false } as ReturnType<typeof refreshPivotTable>,
          () => refreshPivotTable(wb, { sheet: pivot.sheet, pivotIndex: pivot.pivotIndex }),
        );
        if (result.ok) refreshPivotCells(pivot.sheet);
        return;
      }
    }
  }

  host.addEventListener('contextmenu', onContextMenu);
  host.addEventListener('keydown', onHostKey);
  ownerDocument.addEventListener('contextmenu', onDocOutsidePress, true);
  ownerDocument.addEventListener('mousedown', onDocOutsidePress, true);
  ownerDocument.addEventListener('focusin', onDocFocusIn, true);
  ownerDocument.addEventListener('keydown', onDocKey, true);
  ownerWindow?.addEventListener('scroll', onScroll, true);
  ownerWindow?.addEventListener('resize', onResize);
  const unsubscribePolicy = interactionController?.subscribe?.(refreshOpenPermissions) ?? null;

  const detach = ((): void => {
    host.removeEventListener('contextmenu', onContextMenu);
    host.removeEventListener('keydown', onHostKey);
    ownerDocument.removeEventListener('contextmenu', onDocOutsidePress, true);
    ownerDocument.removeEventListener('mousedown', onDocOutsidePress, true);
    ownerDocument.removeEventListener('focusin', onDocFocusIn, true);
    ownerDocument.removeEventListener('keydown', onDocKey, true);
    ownerWindow?.removeEventListener('scroll', onScroll, true);
    ownerWindow?.removeEventListener('resize', onResize);
    unsubscribePolicy?.();
    pivotNumberFormat.detach();
    cancelSubClose();
    root.remove();
    sub.remove();
  }) as ContextMenuHandle;
  detach.setStrings = (next: Strings): void => {
    strings = next;
    hide();
  };
  detach.setOptions = (next?: ContextMenuOptions): void => {
    options = next;
    activeContext = null;
    hide();
  };
  return detach;
}

function setContextMenuItemDisabled(
  button: HTMLButtonElement,
  disabled: boolean,
  reason: string | null,
): void {
  projectDisabledState(button, disabled, reason, {
    datasetKey: 'disabledReason',
  });
}
