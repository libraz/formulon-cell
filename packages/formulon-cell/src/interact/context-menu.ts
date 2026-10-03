import { canExecuteBuiltIn } from '../commands/built-in-command-policy.js';
import {
  clearSelectedContents,
  collectSelectedContentAddresses,
} from '../commands/clear-contents.js';
import type { ClipboardSnapshot } from '../commands/clipboard/snapshot.js';
import { parseTSV } from '../commands/clipboard/tsv.js';
import { clearComment } from '../commands/comment.js';
import { withSelectionFormatOrigin } from '../commands/format.js';
import { type History, recordRepeatableFormatChange } from '../commands/history.js';
import { hyperlinkAt } from '../commands/hyperlinks.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import type {
  CellBatchCommand,
  CellChangeInput,
  InteractionOperation,
  OperationEffect,
  OperationIntent,
  PermissionCode,
  PermissionDecision,
} from '../commands/interaction-policy.js';
import { hiddenInSelection } from '../commands/structure.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { hitZone, layoutForView } from '../render/geometry.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
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
import {
  buildCellEntries,
  buildColEntries,
  buildRowEntries,
  compactMenuEntries,
  type ItemId,
  isBuiltinItemId,
  type MenuKind,
  menuEntryToContextItem,
  PASTE_QUICK_IDS,
  PLAIN_INSERT_IDS,
} from './context-menu-spec.js';
import {
  type ContextMenuStructureContext,
  runContextMenuStructureItem,
} from './context-menu-structure.js';
import { navigationBoundsFor } from './navigation-policy.js';
import { overlayPortalFor } from './overlay-portal.js';
import { clampPanelToViewport, panelSize, viewportSize } from './overlay-position.js';

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
  /** Called when the user clicks "Paste Special…". */
  onPasteSpecial?: () => void;
  /** Called when the user clicks "Define Name…". When omitted the entry is
   *  hidden. */
  onDefineName?: () => void;
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
  /** Opens the active cell's hyperlink after the menu validates its target. */
  onOpenHyperlink?: (url: string) => void;
  /** Called when the user clicks the Add/Remove Watch entry. When omitted the
   *  menu entry is hidden. */
  onToggleWatch?: (addr: Addr) => void;
  /** Returns true when the active cell is currently watched. */
  isWatched?: (addr: Addr) => boolean;
  /** Optional composed menu contract. Omitted preserves the legacy menu. */
  options?: ContextMenuOptions;
  /** Shared authorization/execution service for restricted embeds. */
  interactionController?: ContextMenuInteractionController;
}

const VIEWPORT_PAD = 4;

type MenuTarget = { kind: MenuKind; cell: Addr };

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
  root.className = 'fc-ctxmenu';
  root.setAttribute('role', 'menu');
  root.setAttribute('aria-label', strings.contextMenu.title);
  root.style.display = 'none';
  root.tabIndex = -1;
  const portal = overlayPortalFor(host);
  portal.appendChild(root);

  // Single reusable child panel — the context menu is one level deep.
  const sub = ownerDocument.createElement('div');
  sub.className = 'fc-ctxmenu fc-ctxmenu__sub';
  sub.setAttribute('role', 'menu');
  sub.style.display = 'none';
  sub.tabIndex = -1;
  portal.appendChild(sub);

  let visible = false;
  let menuKind: MenuKind = 'cell';
  let pasteBtnRef: HTMLButtonElement | null = null;
  let activeIndex = -1;
  let focusPanel: 'root' | 'sub' = 'root';
  let restoreFocusEl: HTMLElement | null = null;
  const submenuChildren = new Map<string, RenderMenuEntry[]>();
  let openSub: { id: string; parentBtn: HTMLButtonElement } | null = null;
  let subCloseTimer: ReturnType<typeof setTimeout> | null = null;
  let activeContext: ContextMenuContext | null = null;
  let activeTarget: MenuTarget | null = null;
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
    openSub.parentBtn.setAttribute('aria-expanded', 'false');
    openSub.parentBtn.classList.remove('fc-ctxmenu__item--open');
    openSub = null;
    sub.style.display = 'none';
    sub.replaceChildren();
    if (focusPanel === 'sub') focusPanel = 'root';
  };

  /** Close the submenu after a short grace period — lets the pointer travel
   *  diagonally from a parent row into the child panel without it snapping
   *  shut as it crosses sibling rows. */
  const scheduleSubClose = (): void => {
    if (!openSub) return;
    cancelSubClose();
    subCloseTimer = setTimeout(() => {
      closeSubmenu();
    }, 260);
  };

  const hide = (restoreFocus = false): void => {
    if (!visible) return;
    visible = false;
    closeSubmenu();
    root.style.display = 'none';
    activeIndex = -1;
    focusPanel = 'root';
    const focusTarget = restoreFocusEl;
    restoreFocusEl = null;
    if (restoreFocus) {
      (focusTarget ?? host).focus({ preventScroll: true });
    }
  };

  const activePanel = (): HTMLElement => (focusPanel === 'sub' ? sub : root);

  const panelItems = (panel: HTMLElement): HTMLButtonElement[] =>
    Array.from(panel.querySelectorAll<HTMLButtonElement>('.fc-ctxmenu__item')).filter(
      (btn) => !btn.disabled && btn.getAttribute('aria-disabled') !== 'true',
    );

  const restrictedMenuSafeIds = new Set<ItemId>(['copy', 'paste', 'clear']);
  const navigationUnsupportedIds = new Set<ItemId>([
    'insertCopiedCells',
    'insertCells',
    'deleteCells',
    'rowInsertAbove',
    'rowInsertBelow',
    'rowDelete',
    'colInsertLeft',
    'colInsertRight',
    'colDelete',
  ]);

  const navigationBounded = (): boolean => navigationBoundsFor(store) !== undefined;

  const canPasteToRange = (range: Range | null): boolean => {
    const bounds = navigationBoundsFor(store);
    if (!bounds) return true;
    if (
      !range ||
      range.sheet !== bounds.sheet ||
      range.r0 < bounds.r0 ||
      range.c0 < bounds.c0 ||
      range.r1 > bounds.r1 ||
      range.c1 > bounds.c1
    ) {
      return false;
    }
    if (!interactionController) return true;
    return interactionController.canExecute({
      operation: 'paste',
      origin: 'contextMenu',
      commandId: 'paste',
      effects: [{ kind: 'range', range }],
    }).allowed;
  };

  const hasExplicitPolicy = (): boolean => {
    const policy = interactionController?.policy;
    return (
      interactionController?.restricted === true ||
      (policy !== undefined && Object.keys(policy).length > 0)
    );
  };

  const operationForItem = (id: string): InteractionOperation | null => {
    if (id === 'copy' || id === 'selectAll' || id === 'rowHeight' || id === 'colWidth') {
      return null;
    }
    if (
      id === 'paste' ||
      id === 'pasteSpecial' ||
      id === 'pasteAll' ||
      id === 'pasteFormulas' ||
      id === 'pasteFormulasNumFmt' ||
      id === 'pasteValues' ||
      id === 'pasteValuesNumFmt' ||
      id === 'pasteFormatsOnly' ||
      id === 'pasteTranspose'
    ) {
      return 'paste';
    }
    if (id === 'clear') return 'clear';
    if (
      id === 'cut' ||
      id === 'insertCopiedCells' ||
      id === 'insertCells' ||
      id === 'deleteCells'
    ) {
      return 'moveCells';
    }
    if (id === 'rowInsertAbove' || id === 'rowInsertBelow') return 'insertRows';
    if (id === 'rowDelete') return 'deleteRows';
    if (id === 'colInsertLeft' || id === 'colInsertRight') return 'insertColumns';
    if (id === 'colDelete') return 'deleteColumns';
    if (
      id === 'bold' ||
      id === 'italic' ||
      id === 'underline' ||
      id === 'alignLeft' ||
      id === 'alignCenter' ||
      id === 'alignRight' ||
      id === 'borders' ||
      id === 'formatCells' ||
      id === 'editPhonetic'
    ) {
      return 'format';
    }
    if (id === 'filterClear' || id === 'filterReapply' || id === 'filterByValue') {
      return 'filter';
    }
    if (id === 'sortAsc' || id === 'sortDesc') return 'sort';
    if (id === 'rowHide' || id === 'rowUnhide') return 'resizeRows';
    if (id === 'colHide' || id === 'colUnhide') return 'resizeColumns';
    if (id === 'rowGroup' || id === 'rowUngroup') return 'insertRows';
    if (id === 'colGroup' || id === 'colUngroup') return 'insertColumns';
    if (id === 'insertComment' || id === 'deleteComment') return 'comment';
    if (id === 'insertHyperlink' || id === 'openHyperlink') return 'hyperlink';
    if (id === 'defineName') return 'namedRange';
    if (id === 'toggleWatch') return 'object';
    return 'object';
  };

  const operationEffect = (id: string): OperationEffect => {
    // Row/column structure commands operate on the worksheet axis. Feeding a
    // full row or column range into the authorization materializer would
    // exceed its bounded cell budget before the command's own protection and
    // overflow checks run. Treat these as workbook-structure effects; the
    // navigation gate above still applies to embedded bounded views.
    const structural =
      (id === 'insertCopiedCells' && menuKind !== 'cell') ||
      id === 'rowInsertAbove' ||
      id === 'rowInsertBelow' ||
      id === 'rowDelete' ||
      id === 'colInsertLeft' ||
      id === 'colInsertRight' ||
      id === 'colDelete';
    return structural
      ? { kind: 'workbook' }
      : { kind: 'range', range: { ...store.getState().selection.range } };
  };

  const intentForBuiltIn = (id: string): OperationIntent | null => {
    const operation = operationForItem(id);
    if (!operation) return null;
    return {
      operation,
      origin: 'contextMenu',
      commandId: id,
      effects: [operationEffect(id)],
    };
  };

  const deniedDecision = (code: PermissionCode, reason?: string): PermissionDecision => ({
    allowed: false,
    code,
    ...(reason ? { reason } : {}),
  });

  const decisionForIntent = (intent: OperationIntent): PermissionDecision => {
    if (!interactionController) return { allowed: true };
    return interactionController.canExecute(intent);
  };

  const decisionReason = (decision: PermissionDecision): string | null =>
    decision.allowed ? null : (decision.reason ?? null);

  const decisionForItem = (
    item: ItemId | ContextMenuItem | OperationIntent,
  ): PermissionDecision => {
    if (typeof item === 'object' && 'operation' in item && 'origin' in item) {
      return decisionForIntent(item);
    }
    const itemId = typeof item === 'string' ? item : item.builtIn;
    if (!itemId) {
      const command = typeof item === 'string' ? undefined : item.command;
      if (!command) return { allowed: true };
      return decisionForIntent({
        operation: command.operation,
        origin: 'contextMenu',
        commandId: command.commandId ?? command.type,
        effects: [
          {
            kind: 'cells',
            cells: command.changes.map((change) => change.addr),
            includesFormula: command.changes.some(
              (change) => 'formula' in change && change.formula != null,
            ),
          },
        ],
      });
    }
    if (itemId === 'copy' && interactionController?.policy?.copy === false) {
      return deniedDecision('operationDenied', 'Copy is disabled by the host policy.');
    }
    if (navigationBounded() && navigationUnsupportedIds.has(itemId)) {
      return deniedDecision(
        'unsupported',
        'This structural operation is unavailable with a bounded viewport.',
      );
    }
    if (hasExplicitPolicy() && !restrictedMenuSafeIds.has(itemId)) {
      return deniedDecision(
        'unsupported',
        'This context-menu operation is unavailable in restricted mode.',
      );
    }
    if (hasExplicitPolicy()) return canExecuteBuiltIn(store, itemId, 'contextMenu');
    const intent = intentForBuiltIn(itemId);
    return intent ? decisionForIntent(intent) : { allowed: true };
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
    const items = panelItems(activePanel());
    if (items.length === 0) return;
    activeIndex = (idx + items.length) % items.length;
    items[activeIndex]?.focus();
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
    for (const child of children) appendEntry('sub', sub, child, disabled);
    sub.style.display = 'block';
    sub.style.left = '-9999px';
    sub.style.top = '-9999px';
    const r = parentBtn.getBoundingClientRect();
    const { width: sw } = panelSize(sub);
    const viewport = viewportSize();
    let x = r.right - 2;
    if (x + sw > viewport.width - VIEWPORT_PAD) x = r.left - sw + 2;
    const pos = clampPanelToViewport(sub, x, r.top - 4, { pad: VIEWPORT_PAD });
    sub.style.left = `${pos.x}px`;
    sub.style.top = `${pos.y}px`;
    openSub = { id, parentBtn };
    parentBtn.setAttribute('aria-expanded', 'true');
    parentBtn.classList.add('fc-ctxmenu__item--open');
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
      const btn = createContextSubmenuButton(entry);
      if (source && explicitDisabledSources.has(source)) {
        setContextMenuItemDisabled(btn, true, source.disabledReason ?? null);
      } else if ((source?.builtIn || source?.command) && !decisionForItem(source).allowed) {
        const decision = decisionForItem(source);
        btn.dataset.policyDisabled = 'true';
        setContextMenuItemDisabled(btn, true, decisionReason(decision));
      }
      btn.addEventListener('mouseenter', () => {
        if (btn.disabled) return;
        openSubmenu(entry.id, btn);
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
    const btn = createContextMenuItemButton(entry);
    const source = entry.source;
    if (source) activeItemSources.set(entry.id, source);
    if (disabledIds.has(entry.id)) {
      setContextMenuItemDisabled(btn, true, strings.contextMenu.pasteSpecialRequiresCopiedCells);
    } else if (source && explicitDisabledSources.has(source)) {
      if ((source.builtIn || source.command) && !decisionForItem(source).allowed) {
        btn.dataset.policyDisabled = 'true';
      }
      setContextMenuItemDisabled(btn, true, source.disabledReason ?? null);
    } else if (source?.builtIn && !decisionForItem(source).allowed) {
      const decision = decisionForItem(source);
      btn.dataset.policyDisabled = 'true';
      setContextMenuItemDisabled(btn, true, decisionReason(decision));
    } else if (source?.command && !decisionForItem(source).allowed) {
      const decision = decisionForItem(source);
      btn.dataset.policyDisabled = 'true';
      setContextMenuItemDisabled(btn, true, decisionReason(decision));
    }
    btn.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      if (btn.disabled) return;
      const menuHadFocus =
        ownerDocument.activeElement === btn ||
        root.contains(ownerDocument.activeElement) ||
        sub.contains(ownerDocument.activeElement);
      runEntry(entry);
      hide(false);
      // Context-menu buttons live in the overlay portal. Leaving focus on the
      // now-hidden button prevents the host's keyboard undo/redo listener
      // from seeing the next Cmd/Ctrl+Z. Dialog actions intentionally retain
      // their focus, so only restore the grid host when no dialog owns it.
      if (menuHadFocus && !ownerDocument.activeElement?.closest('[role="dialog"]')) {
        host.focus({ preventScroll: true });
      }
    });
    btn.addEventListener('mouseenter', () => {
      if (panel === 'root') {
        if (openSub && openSub.parentBtn !== btn) scheduleSubClose();
      } else {
        cancelSubClose();
      }
    });
    container.appendChild(btn);
    if (panel === 'root' && entry.id === 'paste') pasteBtnRef = btn;
  };

  const buildMiniToolbar = (
    allowedIds?: ReadonlySet<string>,
    hiddenIds?: ReadonlySet<string>,
  ): HTMLElement => {
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
      const btn = createContextMiniToolbarButton(item);
      const decision = decisionForItem(item.id);
      if (!decision.allowed) {
        btn.dataset.policyDisabled = 'true';
        setContextMenuItemDisabled(btn, true, decisionReason(decision));
      }
      btn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        run(item.id);
        hide(false);
      });
      toolbar.appendChild(btn);
    }

    return toolbar;
  };

  const buildMenu = (kind: MenuKind): void => {
    menuKind = kind;
    root.replaceChildren();
    submenuChildren.clear();
    activeItemSources.clear();
    explicitDisabledSources = new WeakSet<ContextMenuItem>();
    policyDisabledSources = new WeakSet<ContextMenuItem>();
    pasteBtnRef = null;
    const raw =
      kind === 'row'
        ? buildRowEntries(strings)
        : kind === 'col'
          ? buildColEntries(strings)
          : buildCellEntries(strings);
    const activeAddr = store.getState().selection.active;
    const hasCopiedCells = !!store.getState().ui.copyRange;
    const pendingSnapshot = hasCopiedCells ? clipboard.snapshot() : null;
    const pendingCutBand =
      pendingSnapshot?.mode === 'cut' && wholeBandAxisFor(pendingSnapshot) === kind;
    const watched = !!deps.isWatched?.(activeAddr);
    const entries = compactMenuEntries(
      raw
        .filter(
          (e) => !(e.kind === 'item' && e.id === 'insertHyperlink' && !deps.onInsertHyperlink),
        )
        .filter((e) => !(e.kind === 'item' && e.id === 'openHyperlink' && !deps.onOpenHyperlink))
        // A pending copy swaps the plain insert entries for "Insert Copied
        // Cells" — the desktop menus never show both at once.
        .filter((e) => !(e.kind === 'item' && e.id === 'insertCopiedCells' && !hasCopiedCells))
        .filter((e) => !(e.kind === 'item' && hasCopiedCells && PLAIN_INSERT_IDS.includes(e.id)))
        .filter((e) => !(e.kind === 'item' && e.id === 'insertComment' && !deps.onEditComment))
        .filter((e) => !(e.kind === 'item' && e.id === 'toggleWatch' && !deps.onToggleWatch))
        .filter((e) => !(e.kind === 'item' && e.id === 'defineName' && !deps.onDefineName))
        .map((e) => {
          if (e.kind === 'item' && e.id === 'toggleWatch') {
            return {
              ...e,
              label: watched ? strings.contextMenu.removeWatch : strings.contextMenu.addWatch,
            };
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
    if (kind === 'cell') {
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
      hiddenInSelection(s.layout, 'row', s.selection.range.r0, s.selection.range.r1).length > 0;
    const colHidden =
      kind === 'col' &&
      hiddenInSelection(s.layout, 'col', s.selection.range.c0, s.selection.range.c1).length > 0;
    const rowUnhide = root.querySelector<HTMLButtonElement>('[data-fc-action="rowUnhide"]');
    const colUnhide = root.querySelector<HTMLButtonElement>('[data-fc-action="colUnhide"]');
    const openHyperlink = root.querySelector<HTMLButtonElement>('[data-fc-action="openHyperlink"]');
    if (rowUnhide) {
      setContextMenuItemDisabled(rowUnhide, !rowHidden, strings.contextMenu.noHiddenRows);
    }
    if (colUnhide) {
      setContextMenuItemDisabled(colUnhide, !colHidden, strings.contextMenu.noHiddenColumns);
    }
    if (openHyperlink) {
      const target = hyperlinkAt(s, s.selection.active);
      setContextMenuItemDisabled(
        openHyperlink,
        target === null || !isSafeHyperlink(target),
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
    // A fullscreen/dialog resolver can change between the browser's
    // fullscreenchange notification and this key/mouse event. Refresh the
    // owned portal synchronously so the menu is born inside the current focus
    // boundary instead of one event behind.
    overlayPortalFor(host);
    activeTarget = target;
    activeEvent = event;
    activeContext = options ? createContext(target, event, []) : null;
    restoreFocusEl =
      ownerDocument.activeElement instanceof OwnerHTMLElement ? ownerDocument.activeElement : host;
    root.setAttribute('aria-label', strings.contextMenu.title);
    focusPanel = 'root';
    buildMenu(target.kind);
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
    root.style.left = '-9999px';
    root.style.top = '-9999px';
    visible = true;
    const { x, y } = clampPanelToViewport(root, clientX, clientY, { pad: VIEWPORT_PAD });
    root.style.left = `${x}px`;
    root.style.top = `${y}px`;
    focusMenuItem(0);
  };

  /** Resolve which menu flavour to show based on the click target. Header
   *  clicks promote the selection to the whole row/column so the action
   *  inherits a sensible band. */
  const resolveMenuTarget = (e: MouseEvent, updateCellSelection: boolean): MenuTarget => {
    const rect = hitHost.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    const s = store.getState();
    const zone = hitZone(layoutForView(s), s.viewport, x, y, null, { resizeHandles: false });
    const fallback = { kind: 'cell' as const, cell: { ...s.selection.active } };
    if (!zone) return fallback;
    const selectedRanges = [s.selection.range, ...(s.selection.extraRanges ?? [])];
    if (zone.kind === 'row-header' || zone.kind === 'row-resize') {
      const inSel = selectedRanges.some(
        (sel) => zone.row >= sel.r0 && zone.row <= sel.r1 && sel.c0 === 0 && sel.c1 >= 16383,
      );
      if (!inSel && canChangeSelection()) mutators.selectRow(store, zone.row);
      return {
        kind: 'row',
        cell: { ...store.getState().selection.active },
      };
    }
    if (zone.kind === 'col-header' || zone.kind === 'col-resize') {
      const inSel = selectedRanges.some(
        (sel) => zone.col >= sel.c0 && zone.col <= sel.c1 && sel.r0 === 0 && sel.r1 >= 1048575,
      );
      if (!inSel && canChangeSelection()) mutators.selectCol(store, zone.col);
      return {
        kind: 'col',
        cell: { ...store.getState().selection.active },
      };
    }
    if (zone.kind === 'cell') {
      const selected = selectedRanges.find(
        (sel) =>
          zone.row >= sel.r0 && zone.row <= sel.r1 && zone.col >= sel.c0 && zone.col <= sel.c1,
      );
      if (selected?.c0 === 0 && selected.c1 >= 16383) {
        return {
          kind: 'row',
          cell: { sheet: s.selection.active.sheet, row: zone.row, col: zone.col },
        };
      }
      if (selected?.r0 === 0 && selected.r1 >= 1048575) {
        return {
          kind: 'col',
          cell: { sheet: s.selection.active.sheet, row: zone.row, col: zone.col },
        };
      }
      if (!selected && updateCellSelection && canChangeSelection()) {
        const cell = { sheet: s.selection.active.sheet, row: zone.row, col: zone.col };
        mutators.setActive(store, cell);
        return { kind: 'cell', cell };
      }
      return {
        kind: 'cell',
        cell: { sheet: s.selection.active.sheet, row: zone.row, col: zone.col },
      };
    }
    return fallback;
  };

  const isOwnChromeContextTarget = (target: EventTarget | null): boolean =>
    target instanceof OwnerElement &&
    !!target.closest('.fc-host__formulabar, .fc-host__sheetbar, .fc-sheetmenu');

  const canChangeSelection = (): boolean => {
    if (interactionController?.canSelect) return interactionController.canSelect().allowed;
    return interactionController?.policy?.selection !== false;
  };

  const onContextMenu = (e: MouseEvent): void => {
    if (isOwnChromeContextTarget(e.target)) return;
    if (options?.mode === 'disabled') return;
    e.preventDefault();
    const target = resolveMenuTarget(e, options !== undefined);
    show(e.clientX, e.clientY, target, e);
  };

  const onHostKey = (e: KeyboardEvent): void => {
    if (options?.mode === 'disabled') return;
    if (e.key !== 'ContextMenu' && !(e.key === 'F10' && e.shiftKey)) return;
    e.preventDefault();
    const active = store.getState().selection.active;
    const rect = host.getBoundingClientRect();
    show(
      rect.left + Math.min(rect.width || 1, 24),
      rect.top + Math.min(rect.height || 1, 24),
      {
        kind: 'cell',
        cell: { ...active },
      },
      e,
    );
  };

  const insideMenu = (target: EventTarget | null): boolean =>
    target instanceof OwnerNode && (root.contains(target) || sub.contains(target));

  const onDocPointerDown = (e: MouseEvent): void => {
    if (!visible) return;
    if (insideMenu(e.target)) return;
    hide(false);
  };

  const onDocContextMenu = (e: MouseEvent): void => {
    if (!visible) return;
    if (insideMenu(e.target)) return;
    hide(false);
  };

  const onDocKey = (e: KeyboardEvent): void => {
    if (!visible) return;
    if (e.key === 'Escape') {
      e.preventDefault();
      if (openSub) {
        const parent = openSub.parentBtn;
        closeSubmenu();
        focusPanel = 'root';
        parent.focus();
        activeIndex = panelItems(root).indexOf(parent);
      } else {
        hide(true);
      }
    } else if (e.key === 'ArrowDown') {
      e.preventDefault();
      focusMenuItem(activeIndex + 1);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      focusMenuItem(activeIndex - 1);
    } else if (e.key === 'Home') {
      e.preventDefault();
      focusMenuItem(0);
    } else if (e.key === 'End') {
      e.preventDefault();
      focusMenuItem(panelItems(activePanel()).length - 1);
    } else if (e.key === 'ArrowRight') {
      const el = ownerDocument.activeElement;
      if (el instanceof OwnerHTMLButtonElement && el.dataset.fcSubmenu) {
        e.preventDefault();
        openSubmenu(el.dataset.fcSubmenu, el);
        focusPanel = 'sub';
        focusMenuItem(0);
      }
    } else if (e.key === 'ArrowLeft') {
      if (openSub && focusPanel === 'sub') {
        e.preventDefault();
        const parent = openSub.parentBtn;
        closeSubmenu();
        focusPanel = 'root';
        parent.focus();
        activeIndex = panelItems(root).indexOf(parent);
      }
    } else if (e.key === 'Enter' || e.key === ' ') {
      const target = ownerDocument.activeElement;
      if (target instanceof OwnerHTMLButtonElement && insideMenu(target)) {
        e.preventDefault();
        target.click();
      }
    }
  };

  const onScroll = (): void => hide(false);

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
          button.dataset.policyDisabled = 'true';
          setContextMenuItemDisabled(button, true, decisionReason(decision));
        } else if (button.dataset.policyDisabled === 'true') {
          delete button.dataset.policyDisabled;
          if (source && explicitDisabledSources.has(source)) {
            setContextMenuItemDisabled(button, true, source.disabledReason ?? null);
          } else {
            projectDisabledState(button, false, null, { datasetKey: 'disabledReason' });
          }
        }
      }
    }
  };

  sub.addEventListener('mouseenter', cancelSubClose);
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
      if (!interactionController) return;
      try {
        // Commands supplied by a host are still dispatched through the
        // context-menu route. Never trust a caller-provided origin here: an
        // arbitrary origin could bypass the route's policy decision between
        // rendering and activation.
        const command: CellBatchCommand = { ...source.command, origin: 'contextMenu' };
        const result = interactionController.execute(command);
        if (result && typeof result === 'object' && 'then' in result) {
          void Promise.resolve(result).then(() => deps.onAfterCommit?.());
        } else {
          deps.onAfterCommit?.();
        }
      } catch (err) {
        console.warn('formulon-cell: context-menu command failed', err);
      }
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
      denied: interactionController?.policy?.batchDenied,
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

  function run(id: string): void {
    if (!isBuiltinItemId(id)) return;
    if (!decisionForItem(id).allowed) return;
    if (navigationBounded() && navigationUnsupportedIds.has(id)) return;
    const state = store.getState();
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
        if (hasExplicitPolicy()) {
          const addresses = collectSelectedContentAddresses(store, wb);
          if (addresses.length === 0) return;
          executeBatch({
            type: 'cellBatch',
            operation: 'clear',
            origin: 'contextMenu',
            commandId: 'clear',
            changes: addresses.map((addr) => ({ addr, value: { kind: 'blank' as const } })),
            denied: interactionController?.policy?.batchDenied,
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
          if (result.status === 'applied') deps.onAfterCommit?.();
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
        wrapFmt('deleteComment', () => clearComment(store, addr, wb));
        return;
      }
      case 'insertHyperlink': {
        deps.onInsertHyperlink?.();
        return;
      }
      case 'openHyperlink': {
        const target = hyperlinkAt(state, state.selection.active);
        if (!target || !isSafeHyperlink(target)) return;
        deps.onOpenHyperlink?.(target);
        return;
      }
      case 'toggleWatch': {
        deps.onToggleWatch?.(state.selection.active);
        return;
      }
    }
  }

  host.addEventListener('contextmenu', onContextMenu);
  host.addEventListener('keydown', onHostKey);
  ownerDocument.addEventListener('contextmenu', onDocContextMenu, true);
  ownerDocument.addEventListener('mousedown', onDocPointerDown, true);
  ownerDocument.addEventListener('keydown', onDocKey, true);
  ownerWindow?.addEventListener('scroll', onScroll, true);
  const unsubscribePolicy = interactionController?.subscribe?.(refreshOpenPermissions) ?? null;

  const detach = ((): void => {
    host.removeEventListener('contextmenu', onContextMenu);
    host.removeEventListener('keydown', onHostKey);
    ownerDocument.removeEventListener('contextmenu', onDocContextMenu, true);
    ownerDocument.removeEventListener('mousedown', onDocPointerDown, true);
    ownerDocument.removeEventListener('keydown', onDocKey, true);
    ownerWindow?.removeEventListener('scroll', onScroll, true);
    unsubscribePolicy?.();
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

function isSafeHyperlink(url: string): boolean {
  const lower = url.trim().toLowerCase();
  return (
    lower.startsWith('http://') ||
    lower.startsWith('https://') ||
    lower.startsWith('mailto:') ||
    lower.startsWith('tel:')
  );
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
